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
    internal sealed class IVsHierarchy2Calls
    {
        internal static async Task<object[]> GetIVsHierarchyPropertiesAsync(IVsHierarchy hierarchy, JoinableTaskFactory joinableTaskFactory)
        {
            await joinableTaskFactory.SwitchToMainThreadAsync();

            int[] propids = new int[]
            {
                (int)__VSHPROPID.VSHPROPID_Name,
                (int)__VSHPROPID.VSHPROPID_CanBuildFromMemory,
                (int)__VSHPROPID.VSHPROPID_Caption,
                (int)__VSHPROPID.VSHPROPID_DefaultEnableBuildProjectCfg,
                (int)__VSHPROPID.VSHPROPID_DefaultNamespace,
                (int)__VSHPROPID.VSHPROPID_DesignerFunctionVisibility,
                (int)__VSHPROPID.VSHPROPID_EditLabel,
                (int)__VSHPROPID.VSHPROPID_Expandable,
                (int)__VSHPROPID.VSHPROPID_Expanded,
                (int)__VSHPROPID.VSHPROPID_ExtObject,
            };

            object[] values = new object[propids.Length];
            int[] results = new int[propids.Length];

            var hierarchy2 = hierarchy as IVsHierarchy2;
            hierarchy2.GetProperties((uint)VSConstants.VSITEMID.Root, 10, propids, values, results);

            await TaskScheduler.Default;

            return values;
        }
    }
}
