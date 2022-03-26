using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRevalorizacion.Comunes
{
    class ExceptionLogging
    {
        public static void HandleExcepcion(Exception ex, Type type, string CurrentMethodName, string XMLFileName = "")
        {
            ShowStatusBar(ex.Message, (type != null ? type.Name + "." : "") + CurrentMethodName);
            
        }

        internal static void ShowStatusBar(string s_mensaje, string s_ubicacion = "", SAPbouiCOM.BoStatusBarMessageType msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Error, SAPbouiCOM.BoMessageTime msgTypeSec = SAPbouiCOM.BoMessageTime.bmt_Short)
        {
            StringBuilder m_sMsg = new StringBuilder();
            try
            {
                if (!string.IsNullOrEmpty(s_mensaje))
                {
                    m_sMsg.Append("[" + Properties.Resources.NombreAddon + "]");
                    m_sMsg.AppendFormat((s_ubicacion == "" ? "{0}" : " {0}:"), s_ubicacion);
                    m_sMsg.AppendFormat(" {0}", s_mensaje);

                    Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(m_sMsg.ToString(), msgTypeSec, msgType);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("[" + Properties.Resources.NombreAddon + "] CommonsFunctions.cs > ShowStatusBar() | " + s_ubicacion + ": " + ex.Message + " | " + s_mensaje, Properties.Resources.NombreAddon,
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
            finally
            {
                m_sMsg.Length = 0;
            }
        }
    }
}
