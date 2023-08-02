﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExellAddInsLib.MSG
{
    [AttributeUsage(AttributeTargets.Property)]
    public class NonRegisterInUpCellAddresMapAttribute : Attribute
    {
        
    }
    [AttributeUsage(AttributeTargets.Property)]
    public class NonGettinInReflectionAttribute : Attribute
    {

    }
    [AttributeUsage(AttributeTargets.Property)]
    public class DontCloneAttribute : Attribute
    {

    }
}
