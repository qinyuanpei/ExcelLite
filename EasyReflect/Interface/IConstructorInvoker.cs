using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyReflect.Interface
{
    interface IConstructorInvoker
    {
        object Invoke(object[] param);
    }
}
