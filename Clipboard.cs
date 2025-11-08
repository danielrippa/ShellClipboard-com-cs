using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace Shell {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("EC8D72A1-6484-4A5F-B034-8376A7056029")]
  [ProgId("Shell.Clipboard")]
  public class Clipboard {

    public Clipboard() {
    }

    public string GetText() {
      string returnText = null;
      Exception threadEx = null;
      Thread staThread = new Thread(
        delegate () {
          try {
            returnText = System.Windows.Forms.Clipboard.GetText();
          } catch (Exception ex) {
            threadEx = ex;
          }
        });
      staThread.SetApartmentState(ApartmentState.STA);
      staThread.Start();
      staThread.Join();

      if (threadEx != null) {
        return Serialize(new { error = threadEx.Message });
      }

      return Serialize(new { value = returnText });
    }

    public string SetText(string text) {
      Exception threadEx = null;
      Thread staThread = new Thread(
        delegate () {
          try {
            System.Windows.Forms.Clipboard.SetText(text);
          } catch (Exception ex) {
            threadEx = ex;
          }
        });
      staThread.SetApartmentState(ApartmentState.STA);
      staThread.Start();
      staThread.Join();

      if (threadEx != null) {
        return Serialize(new { error = threadEx.Message });
      }

      return Serialize(new { success = true });
    }

    private string Serialize(object obj) {
      var serializer = new JavaScriptSerializer();
      return serializer.Serialize(obj);
    }
  }
}
