using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyReflect.Interface
{
    interface IMethodInvoker
    {
        object Invoke(object target, object[] param);
    }
}
