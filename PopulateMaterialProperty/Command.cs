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
      //IEnumerable<Element> col
        = new FilteredElementCollector( doc )
          .WhereElementIsNotElementType()
          .WherePasses( ExportParameters.GetFilter() );
          //.Where<Element>( e => ElementId.InvalidElementId == e.GroupId );

      //Element e1 = col.FirstElement();
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

          doc.Regenerate();

          t.Commit();
        }
      }

      Definition def = ExportParameters.GetDefinition( e1 );

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
            p.Set( material_name );
          }
        }
        tx.Commit();
      }
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
