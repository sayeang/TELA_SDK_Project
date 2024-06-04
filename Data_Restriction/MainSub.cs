
//  SAP MANAGE UI API 2007 SDK Sample
//****************************************************************************
//
//  File:      HelloWorld.cs
//
//  Copyright (c) SAP MANAGE
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
//****************************************************************************
// BEFORE STARTING:
// 1. Add reference to the "SAP Business One UI API"
// 2. Insert the development connection string to the "Command line argument"
//-----------------------------------------------------------------
// 1.
//    a. Project->Add Reference...
//    b. select the "SAP Business One UI API 2007" From the COM folder
//
// 2.
//     a. Project->Properties...
//     b. choose Configuration Properties folder (place the arrow on Debugging)
//     c. place the following connection string in the 'Command line arguments' field
// 0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056
//
//**************************************************************************************************

using Microsoft.VisualBasic;
using System;
namespace SysForm {
    sealed public class MainSub { 
        
        public static void Main() { 
            
            SystemForm SBOSysForm = null; 
            
            SBOSysForm = new SystemForm(); 
            
            //  Starting the Application
            System.Windows.Forms.Application.Run(); 
            
        }    
    }   
} 
