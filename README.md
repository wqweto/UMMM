## Unattended Make My Manifest
VB6 Manifest Creation Tool

### Registration-Free COM

[Registration-Free Activation of COM Components: A Walkthrough](http://msdn.microsoft.com/en-us/library/ms973913.aspx)

UMMM is a tool that can be using in automated builds to create manifests for registration-free COM activation. The tool uses an ini file that describes referenced COM/.NET components. All the classes and interfaces from a referenced component are extracted and included in the application (exe file) manifest at build time.

### Sample usage

Here is a sample multi-project solution with references between VB6 projects and references to external components.

    [Main App (exe)] ---> [Component A (dll)] ---> [TreeView Component (ocx)]
           |                     |
           |                     v
           +------------> [Component B (dll)] ---> [Component C (dll)]
           |
           +------------> [CLR Component (.Net dll)]

UMMM can be used to produce manifest for Main App (exe) that references Component A, Component B, Component C, CLR Component and TreeView Component all at once. No manifest for any other project/component is needed to get registration-free COM working.

Here is a sample ini file for the sample solution above:

    Identity C:\Builds\MainApp.exe MyCompany.MyApp "My Application 1.0"
    File ComponentA.dll
	File ComponentB.dll
	File ComponentC.dll
	File TreeView.ocx
	Dependency ClrComponent.dll /u
	Dependency ComCtl
	
`Identity` command is needed to produce `assemblyIdentity` tag in the output manifest. The  parameters of the command are `<exe_file>` and optional `[name]` and `[description]`. The path of the `<exe_file>` paramter sets the base path of the application, so that relative paths in next commands are based on this application path.

### Referencing COM components

`File` command produces `file` tags in the otuput manifest. This is how all referenced COM components are listed in the ini file. The parameters of the command are `<file_name>` and optional `[interfaces]` and  `[classes_filter]`. If `<file_name>` uses relative path then base application path from `Identity` command is used. The `[interfaces]` parameter is needed only if a multi-threading application is passing object references between apartments. The `[classes_filter]` uses Like operator to filter referenced component coclass names to be included in the output manifest.

### Referencing .Net components

`Dependency` produces `dependency` tag in output manifest. This is how .Net components are listed in the ini file. The parameters of the command are `<assembly_file>` and optional `[version]` and `[/update]`. If `<assembly_file>` uses relative path, base path from `Identity` is used. Optional `[/update]` can be used to embed assembly manifest in the assembly file.

UMMM uses `mt.exe` to generate (temporary) manifest from the assembly file and to extract `assemblyIdentity` tag from this manifest. Passing `[/update]` parameter instructs UMMM to embed this temporary manifest in the assembly file using `mt.exe` again. Manifest Tool (`mt.exe`) is installed with Visual Studio in C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin. Make sure `mt.exe` is accessible by `ummm.exe` -- copy it in the same folder or set `PATH` to include `mt.exe`'s folder.

### Referencing Common Controls v6.0

`Dependency` has a second form that accepts `<lib_name>` as first parameter. Built-in `comctl`, `vc90crt`, `vc90mfc` are recognized and correct `publicKeyToken`s are outputed under `dependentAssembly` tag.

### List of commands

Command names are case-insensitive.

#### Identity

Appends `assemblyIdentity` and `description` tags.

    Parameters       <exe_file> [name] [description]
       exe_file      file name can be quoted if containing spaces. The containing folder 
                     of the executable sets base path for relative file names
       name          (optional) assembly name. Defaults to MyAssembly
       description   (optional) description of assembly

#### Dependency

Appends `dependency` tag for referencing dependent assemblies like Common Controls 6.0, C run-time or MFC.

    Parameters       {<lib_name>|<assembly_file>} [version] [/update]
      lib_name       one of { comctl, vc90crt, vc90mfc }
      assembly_file  file name of .NET DLL exporting COM classes
      version        (optional) required assembly version. Multiple version of vc90crt can
                     be required by a single manifest
      /update        (optional) updates assembly_file assembly manifest. Spawns mt.exe

#### File

Appends `file` tag and collects information about coclasses and interfaces exposed by the referenced COM component typelib.

    Parameters       <file_name> [interfaces] [classes_filter]
      file_name      file containing typelib. Can be relative to base path
      interfaces     (optional) pipe (|) separated interfaces with or w/o leading 
                     underscore
      classes_filter (optional) pipe (|) separated filter for coclasses in file

#### Interface

Appends `comInterfaceExternalProxyStub` tag for inter-thread marshaling of interfaces.

    Parameters       <file_name> <interfaces>
      file_name      file containing typelib. Can be relative to base path
      interfaces     pipe (|) separated interfaces with or w/o leading underscore

#### TrustInfo

Appends `trustInfo` tag for UAC user-rights elevation on Vista and above.

    Parameters       [level] [uiaccess]
      level          (optional) one of { 1, 2, 3 } corresponding to { asInvoker, 
                     highestAvailable, requireAdministrator }. Default is 1
      uiaccess       (optional) true/false or 0/1. Allows application to gain access to 
                     the protected system UI. Default is 0

#### DpiAware

Appends `dpiAware` tag for custom DPI aware applications.

    Parameters       [on_off] [per_monitor]
      on_off         (optional) true/false or 0/1. Default is 0
      per_monitor    (optional) win81 per-monitor DPI awareness. Default is 0

#### SupportedOS

Appends `supportedOS` tag.

    Parameters       <os_type> [os_type #2] [os_type #3] ...
      os_type        one of { vista, win7, win8, win81, win10 } or raw GUID as specified
                     by Microsoft. Multiple OSes can be included in a manifest
