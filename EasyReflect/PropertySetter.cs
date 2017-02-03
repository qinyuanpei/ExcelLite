using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyReflect.Interface;
using System.Reflection;

namespace EasyReflect
{
    class PropertySetter<T,V>
    {
        
        public PropertySetter(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
                throw new ArgumentException("Please confirm whether propertyInfo is null!");
            if (!propertyInfo.CanWrite)
                throw new ArgumentException("Please confirm whether propertyInfo is support write!");

            MethodInfo methodInfo = propertyInfo.GetSetMethod(true);

        }
    }
}
