using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSIXProject1
{
    internal sealed class IVsHierarchyCalls
    {
        internal static async Task<object> GetIVsHierarchyPropertyAsync(IVsHierarchy hierarchy, int propId, JoinableTaskFactory joinableTaskFactory)
        {
            await joinableTaskFactory.SwitchToMainThreadAsync();

            int hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, propId, out object value);

            await TaskScheduler.Default;

            return value;
        }

        internal static async Task<object[]> GetIVsHierarchyPropertiesAsync(IVsHierarchy hierarchy, JoinableTaskFactory joinableTaskFactory)
        {
            await joinableTaskFactory.SwitchToMainThreadAsync();

            var values = new object[10];

            int hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out object value);
            values[0] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_CanBuildFromMemory, out value);
            values[1] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out value);
            values[2] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_DefaultEnableBuildProjectCfg, out value);
            values[3] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_DefaultNamespace, out value);
            values[4] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_DesignerFunctionVisibility, out value);
            values[5] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_EditLabel, out value);
            values[6] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Expandable, out value);
            values[7] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Expanded, out value);
            values[8] = value;

            hr = hierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_ExtObject, out value);
            values[9] = value;

            await TaskScheduler.Default;

            return values;
        }
    }
}
