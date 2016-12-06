#define NEED_SEPARATE_COMMAND_TO_CREATE

#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace PopulateMaterialProperty
{
  [Transaction( TransactionMode.Manual )]
  public class Command : IExternalCommand
  {
    /// <summary>
    /// Implement a dictionary to map unwanted source 
    /// material names to the cleaned-ip desireable 
    /// target material names.
    /// </summary>
    static Dictionary<string, string> _map = null;

    /// <summary>
    /// Define the material name string mapping pairs.
    /// </summary>
    static string[] _map_entry_pairs = new string[]
    {
      "Masonry - Brick-BLDG 1-King Facade- 3Levelup-13", "Brick",
      "Masonry - Brick", "Brick",
      "Concrete - Precast Concrete", "Concrete",
      "BLD4-Graybrick-Facade", "Brick",
      "Concrete - Cast-in-Place Concrete", "Concrete",
      "brick-Facade-level 1 to 3", "Brick",
      "Default Wall", "Concrete",
      "Metal - Stud Layer", "Steel",
      "Wood - Flooring-ElevatorCore", "Wood",
      "Metal", "Steel",
      "Masonry - Concrete Masonry Units", "Brick",
      "Metal - Aluminum_Corregated", "Aluminum",
      "Sash 2", "Steel",
      "Sill", "Steel",
      "Metal - Aluminum", "Aluminum",
      "wood Panel", "Wood",
      "BLDG 1-Wood - Flooring", "Wood",
      "BLDG 1,2,3-Wood - Flooring-plywood", "Wood",
      "Structure - Wood Joist/Rafter Layer", "Wood",
      "Default Floor", "Wood",
      "Metal - Steel", "Steel",
      "Wood - Dimensional Lumber", "Wood",
      "Metal - Steel - ASTM A992", "Steel",
      "Extruded aluminum headrail", "Aluminum",
      "Metal - Chrome", "Chrome",
      "Wood - Cherry", "Wood",
      "Wood_RoofDeck", "Wood",
      "Mechanical Equipment", "Steel",
      "Ducts", "Steel",
      "Duct Fittings", "Steel",
      "Electrical Fixtures", "Copper"
    };

    /// <summary>
    /// Map certain undesired complex material 
    /// names to something more useful.
    /// </summary>
    string Map( string material_name )
    {
      if( null == _map )
      {
        _map = new Dictionary<string, string>();
        int n = _map_entry_pairs.Length;
        Debug.Assert( 0 == n % 2, "expected an even "
          + " number of string entires to form from-to "
          + " mapping pairs" );
        for(int i=0; i < n; ++i,++i )
        {
          _map.Add( _map_entry_pairs[i], _map_entry_pairs[i + 1] );
        }
      }
      return _map.ContainsKey( material_name )
        ? _map[material_name]
        : material_name;
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      FilteredElementCollector col
        = new FilteredElementCollector( doc )
          .WhereElementIsNotElementType()
          .WherePasses( ExportParameters.GetFilter() );

      Element e1 = col.First<Element>();

      ExportParameters exportParameters
        = new ExportParameters( e1 );

      if( !exportParameters.IsValid )
      {
        //Util.ErrorMsg( "Please initialise the shared "
        //  + "parameters before launching this command." );

        //return Result.Failed;

        using( Transaction t = new Transaction( doc ) )
        {
          t.Start( "Create Shared Parameter" );

          ExportParameters.Create( doc );

          // Need to regenerate after creating shared 
          // parameter before starting to populate it.
          // Or not, after all?

          //doc.Regenerate();

          t.Commit();
        }
      }

      Definition def = ExportParameters.GetDefinition( e1 );

      int n = 0;

      using( Transaction tx = new Transaction( doc ) )
      {
        tx.Start( "Populate Forge Material Shared Parameter" );

        string material_name;

        foreach( Element e in col )
        {
          material_name = string.Empty;

          Category cat = e.Category;

          Debug.Assert( null != cat, "expected all "
            + "elements of interest to have a valid "
            + "category" );

          Debug.Print( "{0} '{1}' '{2}'",
            e.Id.IntegerValue, e.Name,
            ( null == cat ? "<null>" : cat.Name ) );

          if( null != cat && cat.HasMaterialQuantities )
          {
            ICollection<ElementId> materialIds = e.GetMaterialIds( false );
            foreach( ElementId id in materialIds )
            {
              double area = e.GetMaterialArea( id, false );
              if( 0 < area )
              {
                material_name = doc.GetElement( id ).Name;

                Debug.Print( "{0} '{1}' '{2}' '{3}'",
                  e.Id.IntegerValue, e.Name,
                  cat.Name, material_name );

                break;

                // Some elements throw an exception when 
                // setting the parameter. Why? Maybe due
                // to belonging to a group?

                //try
                //{
                //  e.get_Parameter( def ).Set( materialName );
                //}
                //catch( Exception ex )
                //{
                //  Debug.Print( ex.Message );
                //}

                // Added null == GroupId to the 
                // filtered element collector. 
                // That did not help.

                // Added check for read-only parameter.

                //Parameter p = e.get_Parameter( def );

                //Debug.Assert( null != p, "expected all "
                //  + "elements of interest to be equipped "
                //  + "with a shared parameter" );

                //if( !p.IsReadOnly )
                //{
                //  p.Set( materialName );
                //  done = true;
                //  break;
                //}
              }
            }
          }
          if( 0 == material_name.Length )
          {
            material_name = cat.Name;
          }

          Parameter p = e.get_Parameter( def );

          Debug.Assert( null != p, "expected all "
            + "elements of interest to be equipped "
            + "with a shared parameter" );

          if( !p.IsReadOnly )
          {
            material_name = Map( material_name );
            p.Set( material_name );
            ++n;
          }
        }
        tx.Commit();
      }

      Util.InfoMsg( string.Format(
        "Populated {0} Forge material parameter{1}.",
        n, ( 1 == n ? "" : "s" ) ) );

      return Result.Succeeded;
    }
  }
  #region CmdCreateSharedParameters
#if NEED_SEPARATE_COMMAND_TO_CREATE
  [Transaction( TransactionMode.Manual )]
  public class CmdCreateSharedParameters
    : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      Document doc = uiapp.ActiveUIDocument.Document;

      using( Transaction t = new Transaction( doc ) )
      {
        t.Start( "Create Shared Parameter" );
        ExportParameters.Create( doc );
        t.Commit();
      }
      return Result.Succeeded;
    }
  }
#endif // NEED_SEPARATE_COMMAND_TO_CREATE
  #endregion // CmdCreateSharedParameters
}
