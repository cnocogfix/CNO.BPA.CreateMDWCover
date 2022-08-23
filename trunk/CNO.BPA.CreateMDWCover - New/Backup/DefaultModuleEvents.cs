using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Emc.InputAccel.QuickModule.ClientScriptingInterface;
using Emc.InputAccel.ScriptEngine.Scripting;

namespace CNO.BPA.CreateMDWCover
{
    public class DefaultModuleEvents : IDefaultModuleEvents
    {
        public void ModuleStart(IApplication arg)
        {
        }

        public void ModuleFinish()
        {
        }

        public void GotServerConnection(IServerInformation serverInformation)
        {
        }

        public void LostServerConnection(String serverHostName)
        {
        }
    }
}
