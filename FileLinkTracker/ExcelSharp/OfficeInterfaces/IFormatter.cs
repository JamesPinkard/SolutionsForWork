﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp
{
    public interface IFormatter : ISheetTool
    {
        void Format();
    }
}
