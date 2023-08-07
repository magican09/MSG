using System;

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
    public class DoNotCloneAttribute : Attribute
    {

    }
}
