using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyReflect.Interface
{
    interface IValueSetter
    {
        void SetValue(object target, object value);
    }
}
