﻿using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("InspectorEvents")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("")]
[assembly: AssemblyCopyright("")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("B1946E21-B49B-47DA-B891-E72560F150D7")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

internal static class DesignTimeConstants {
    internal const string Version = "10.0.0.0";
    internal const string DesignerAssembly = "Microsoft.VisualStudio.Tools.Office.Designer, Version=" + Version + ", Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    internal const string TypeCodeDomSerializer = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design";
    internal const string RibbonTypeSerializer = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, " + DesignerAssembly;
    internal const string FormRegionTypeCodeDomSerializer = "Microsoft.VisualStudio.OfficeTools.Outlook2007.Design.FormRegionControlTypeCodeDomSerializer, " + DesignerAssembly;
}
