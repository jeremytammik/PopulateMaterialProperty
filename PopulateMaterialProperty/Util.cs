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

    /// <summary>
    /// Display a short big message.
    /// </summary>
    public static void InfoMsg( string msg )
    {
      Debug.Print( msg );
      TaskDialog.Show( Caption, msg );
    }

    /// <summary>
    /// Display a longer message in smaller font.
    /// </summary>
    public static void InfoMsg2(
      string instruction,
      string msg,
      bool prompt = true )
    {
      Debug.Print( string.Format( "{0}: {1}",
        instruction, msg ) );

      if( prompt )
      {
        TaskDialog dlg = new TaskDialog( Caption );
        dlg.MainInstruction = instruction;
        dlg.MainContent = msg;
        dlg.Show();
      }
    }

    /// <summary>
    /// Display an error message.
    /// </summary>
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
