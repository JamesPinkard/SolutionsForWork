﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reader.ProCode
{
    public interface IFieldManager
    {
        List<string> AvailableFields { get; set; }
        List<string> SelectedFields { get; set; }
    }
}
