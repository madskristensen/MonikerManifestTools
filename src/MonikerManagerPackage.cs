using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace MonikerManifestTools
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.guidPackageString)]
    [ProvideUIContextRule(PackageGuids.guidVsixProjectString,
        name: "Only VSIX projects",
        expression: "VSIX project",
        termNames: new[] { "VSIX project" },
        termValues: new[] { "ActiveProjectFlavor:{82b43b9b-a64c-4715-b499-d71e9ca2bd60}" })]
    public sealed class MonikerManagerPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();
            await CreateManifestCommand.InitializeAsync(this);
        }
    }
}
