﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelSharp
{
    public interface IOfficeWriter : ISheetTool
    {
        void Write();
    }
}
