﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ExcelSerialPortListener {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("ExcelSerialPortListener.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to ExcelSerialPortListener.
        /// </summary>
        internal static string ErrorCaption {
            get {
                return ResourceManager.GetString("ErrorCaption", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Excel is not running or requested spreadsheet is not open, exiting now.
        /// </summary>
        internal static string ExcelIsNotRunning {
            get {
                return ResourceManager.GetString("ExcelIsNotRunning", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Excel is not running, please open Excel with the appropriate spreadsheet..
        /// </summary>
        internal static string ExcelIsNotRunningPleaseOpenExcel {
            get {
                return ResourceManager.GetString("ExcelIsNotRunningPleaseOpenExcel", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Expected 3 arguments: WorkbookName, WorkSheetName, Range.
        /// </summary>
        internal static string Expected3Arguments {
            get {
                return ResourceManager.GetString("Expected3Arguments", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Failed to open serial port connection.
        /// </summary>
        internal static string FailedToOpenSerialPortConnection {
            get {
                return ResourceManager.GetString("FailedToOpenSerialPortConnection", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Failed to write to spreadsheet.
        /// </summary>
        internal static string FailedToWriteToSpreadsheet {
            get {
                return ResourceManager.GetString("FailedToWriteToSpreadsheet", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to got Print command..
        /// </summary>
        internal static string GotPrintCommand {
            get {
                return ResourceManager.GetString("GotPrintCommand", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Received Response: {0}.
        /// </summary>
        internal static string ReceivedResponse0 {
            get {
                return ResourceManager.GetString("ReceivedResponse0", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Saw pressed key!.
        /// </summary>
        internal static string SawPressedKey {
            get {
                return ResourceManager.GetString("SawPressedKey", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Timed Out.
        /// </summary>
        internal static string TimedOut {
            get {
                return ResourceManager.GetString("TimedOut", resourceCulture);
            }
        }
    }
}
