#region Namespaces
using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
#endregion // Namespaces

namespace PopulateMaterialProperty
{
  /// <summary>
  /// Shared parameters to define element 
  /// material for Forge export.
  /// </summary>
  class ExportParameters
  {
    /// <summary>
    /// Which element categories are we going to populate?
    /// </summary>
    public static BuiltInCategory[] Bics
      = new BuiltInCategory[] {
          BuiltInCategory.OST_Walls,
          BuiltInCategory.OST_Floors,
          BuiltInCategory.OST_Ceilings,
          BuiltInCategory.OST_DuctFitting,
          BuiltInCategory.OST_DuctCurves,
          BuiltInCategory.OST_LightingFixtures,
          BuiltInCategory.OST_ElectricalFixtures,
          BuiltInCategory.OST_MechanicalEquipment
    };

    /// <summary>
    /// Return an element filter for the
    /// element categories we are interested in.
    /// </summary>
    public static ElementFilter GetFilter()
    {
      IList<ElementFilter> a = new List<ElementFilter>(
        Bics.Length );

      foreach( BuiltInCategory bic in Bics )
      {
        a.Add( new ElementCategoryFilter( bic ) );
      }

      return new LogicalOrFilter( a );
    }

    /// <summary>
    /// Define the user visible export 
    /// history shared parameter names.
    /// </summary>
    const string _forge_material = "ForgeMaterial";

    //Guid _guid_forge_material;

    /// <summary>
    /// Store the shared parameter definitions.
    /// </summary>
    Definition _definition_forge_material = null;

    Document _doc = null;
    List<ElementId> _ids = null;

    /// <summary>
    /// Return the parameter definition from
    /// the given element and parameter name.
    /// </summary>
    static Definition GetDefinition(
      Element e,
      string parameter_name )
    {
      IList<Parameter> ps = e.GetParameters(
        parameter_name );

      int n = ps.Count;

      Debug.Assert( 1 >= n,
        "expected maximum one shared parameters "
        + "named '" + parameter_name + "'" );

      Definition d = ( 0 == n )
        ? null
        : ps[0].Definition;

      return d;
    }

    /// <summary>
    /// Return the one and only parameter definition
    /// that was created by this class.
    /// </summary>
    public static Definition GetDefinition( Element e )
    {
      return GetDefinition( e, _forge_material );
    }

    /// <summary>
    /// Initialise the shared parameter definitions
    /// from a given sample element.
    /// </summary>
    public ExportParameters( Element e )
    {
      _definition_forge_material = GetDefinition(
        e, _forge_material );

      if( IsValid )
      {
        _doc = e.Document;
        _ids = new List<ElementId>();
      }
    }

    /// <summary>
    /// Check whether all parameter definitions were 
    /// successfully initialised.
    /// </summary>
    public bool IsValid
    {
      get
      {
        return null != _definition_forge_material;
      }
    }

    /// <summary>
    /// Create the shared parameters.
    /// </summary>
    public static void Create( Document doc )
    {
      /// <summary>
      /// Shared parameters filename; used only in case
      /// none is set and we need to create the export
      /// history shared parameters.
      /// </summary>
      const string _shared_parameters_filename
        = "shared_parameters.txt";

      const string _definition_group_name
        = "ForgeMaterial";

      Application app = doc.Application;

      // Retrieve shared parameter file name

      string sharedParamsFileName
        = app.SharedParametersFilename;

      if( null == sharedParamsFileName
        || 0 == sharedParamsFileName.Length )
      {
        string path = Path.GetTempPath();

        path = Path.Combine( path,
          _shared_parameters_filename );

        StreamWriter stream;
        stream = new StreamWriter( path );
        stream.Close();

        app.SharedParametersFilename = path;

        sharedParamsFileName
          = app.SharedParametersFilename;
      }

      // Retrieve shared parameter file object

      DefinitionFile f
        = app.OpenSharedParameterFile();

      //using( Transaction t = new Transaction( doc ) )
      //{
      //  t.Start( "Create Shared Parameter" );

      // Create the category set for binding

      CategorySet catSet = app.Create.NewCategorySet();

      foreach( BuiltInCategory bic in Bics )
      {
        Category cat = doc.Settings.Categories.get_Item( bic );

        catSet.Insert( cat );
      }

      Binding binding = app.Create.NewInstanceBinding(
        catSet );

      // Retrieve or create shared parameter group

      DefinitionGroup group
        = f.Groups.get_Item( _definition_group_name )
        ?? f.Groups.Create( _definition_group_name );

      // Retrieve or create the parameters;
      // we could check if they are already bound, 
      // but it looks like Insert will just ignore 
      // them in that case.

      Definition definition = group.Definitions.get_Item( _forge_material );

      if( null == definition )
      {
        ExternalDefinitionCreationOptions opt 
          = new ExternalDefinitionCreationOptions(
            _forge_material, ParameterType.Text );

        definition = group.Definitions.Create( opt );
      }

      doc.ParameterBindings.Insert( definition, binding,
        BuiltInParameterGroup.PG_GENERAL );

      //  t.Commit();
      //}
    }
  }
}
