using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyReflect.Interface
{
    interface IValueGetter
    {
        object GetValue(object target);
    }
}
