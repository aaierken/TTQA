using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


/// <summary>
/// Enumeration for intBrowser
/// </summary>
public enum eBrowser
{
    IE,
    FireFox,
    Chrome
}

/// <summary>
/// Enumeration for CtrlType
/// </summary>
public enum CtrlType
{
    ClassName,
    
    CssSelector,
    
    ID,
    
    LinkText,
    
    Name,

    Text,

    PartialLinkText,
    
    TagName,

    XPath

}

/// <summary>
/// Enumeration for CtrlClass
/// </summary>
public enum CtrlClass
{
    XPath,
    
    CssSelector,
    
    ID,
    
    LinkText,
    
    Name,

    Text,

    Link,
    
    CheckBox,

    RadioButton,

    Button,

    DropDownList,

    TextBox,

    SingleClick,

    DoubleClick
}


/// <summary>
/// 
/// </summary>
public enum CheckStatus
{
    Check,
    UnCheck,
    JustClick
}

/// <summary>
/// 
/// </summary>
public enum BtnClick
{
    OK,
    Cancel
}

/// <summary>
/// 
/// </summary>
public enum TableGrid
{
    Table,

    Grid
}


    public static class GlobalVariables
    {
        /// <summary>Email username vairable</summary>
        /// <summary>SendEmailUserName<value>mnm704</value></summary>
        public static string SendEmailUserName = "mnm704";
        
        /// <summary>Email password variable</summary>
        /// <summary>SendEmailUserPassword<value>704ebay704</value></summary>
        public static string SendEmailUserPassword = "704ebay704";

        ///<summary>URL for QA Project Site</summary>
        ///<summary>qaProjectSiteUrl<value>https://qa.google.com/SignOn.aspx</value> </summary>
        public static string qaProjectSiteUrl = @"https://qa.google.com/SignOn.aspx";

        ///<summary>Running IE Explorer Process</summary>
        ///<summary>ieProcess<value>"IEEXPLORER"</value> </summary>
        public static string ieProcess = "IEXPLORE";

        ///<summary>defaultWaitTime is 30 seconds</summary>
        ///<summary>defaultWaitTime<value>30</value> </summary>
        public static int defaultWaitTime = 30;

        ///<summary>Website file type control use in artifact upload method</summary>
        ///<summary>iEBrowserFileTypeCtrl<value>"file"</value> </summary>
        public static string iEBrowserFileTypeCtrl = "file";

        
    }

