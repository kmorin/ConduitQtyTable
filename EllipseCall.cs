/****************************** Module Header ******************************\
Module Name:  MyMLeaderJig.cs
Project:      ConduitQtyTable
Copyright (c) 2012 Kyle T. Morin.

<Description of the file>

All rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System.Runtime.InteropServices;

namespace ConduitQtyTable
{
    class EllipseCall
    {
        [DllImport("acad.exe", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedCmd")]

        private static extern int acedCmd(System.IntPtr vlist);

        //[CommandMethod("EllipseJigWait")]

        public static bool EllipseJigWait()
        {

            ResultBuffer rb = new ResultBuffer();
            //RTSTR = 5005
            rb.Add(new TypedValue(5005, "_.ELLIPSE"));
            //Start the command
            acedCmd(rb.UnmanagedObject);

            bool quit = false;
            //loop round while the command is active
            while (!quit)
            {
                //see what commands are active
                string cmdNames = (string)Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CMDNAMES");
                //if the command is active
                if (cmdNames.ToUpper().IndexOf("ELLIPSE") >= 0)
                {
                    rb = new ResultBuffer();
                    //RTSTR = 5005 - send user pause
                    rb.Add(new TypedValue(5005, "\\"));
                    acedCmd(rb.UnmanagedObject);
                }
                else
                {
                    quit = true;
                }
            }
            return quit;
        }
    }
}
