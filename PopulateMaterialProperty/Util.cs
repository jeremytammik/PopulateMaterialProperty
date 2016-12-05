#region Namespaces
using System;
using System.Diagnostics;
using Autodesk.Revit.UI;
#endregion // Namespaces

namespace PopulateMaterialProperty
{
  class Util
  {
    public const string Caption = "Populate Material Parameter";

    public static void InfoMsg( string msg )
    {
      Debug.Print( msg );
      TaskDialog.Show( Caption, msg );
    }

    public static void ErrorMsg( string msg )
    {
      Debug.Print( msg );
      TaskDialog dlg = new TaskDialog( Caption );
      dlg.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
      dlg.MainInstruction = msg;
      dlg.Show();
    }

  }
}
