using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("SharedAddInEvents")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("SharedAddInEvents")]
[assembly: AssemblyCopyright("")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

[assembly: AssemblyVersion("1.0.0.0")]

[assembly: AssemblyDelaySign(false)]
[assembly: AssemblyKeyFile("")]
[assembly: AssemblyKeyName("")]

// Make sure we have an explicit LIBID
[assembly: Guid("45C1778F-EFB9-4d93-81EF-EE14011C133B")]

// Also note that Register for COM interop registers the ComVisible
// interfaces and classes, and the typelib, but it doesn't associate
// the classes with the typelib. So, we need to add an extra reg
// key for the typelib under the CLSID for the AddInUtilities object:
// [HKEY_CLASSES_ROOT\<Wow6432Node>\CLSID\{F743C9A0-DDEF-49D5-AEAA-2E6798814C23}\TypeLib]
// @="{45C1778F-EFB9-4d93-81EF-EE14011C133B}"

