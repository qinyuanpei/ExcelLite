using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using EasyReflect.Interface;

namespace EasyReflect.Factory
{
    public static class ReflectFactory
    {
        private static readonly Dictionary<PropertyInfo, IValueSetter> m_SetterDict =
            new Dictionary<PropertyInfo, IValueSetter>();

        private static readonly Dictionary<PropertyInfo, IValueGetter> m_GetterDict =
            new Dictionary<PropertyInfo, IValueGetter>();

        private static readonly Dictionary<ConstructorInfo, IConstructorInvoker> m_ConstructorInvokerDict =
            new Dictionary<ConstructorInfo, IConstructorInvoker>();

        private static readonly Dictionary<MethodInfo, IMethodInvoker> m_MethodInvokerDict =
            new Dictionary<MethodInfo, IMethodInvoker>();
        
        public static IValueSetter GetPropertySetter(PropertyInfo propertyInfo)
        {
            if (!m_SetterDict.ContainsKey(propertyInfo)){
                IValueSetter setter = CreatePropertySetter(propertyInfo);
                m_SetterDict.Add(propertyInfo, setter);
                return setter;
            }

            return m_SetterDict[propertyInfo];
        }

        public static IValueGetter GetPropertyGetter(PropertyInfo propertyInfo)
        {
            if (!m_GetterDict.ContainsKey(propertyInfo)){
                IValueGetter getter = CreatePropertyGetter(propertyInfo);
                m_GetterDict.Add(propertyInfo, getter);
                return getter;
            }

            return m_GetterDict[propertyInfo];
        }

        public static IConstructorInvoker GetConstructorInvoker(ConstructorInfo constructorInfo)
        {
            if (!m_ConstructorInvokerDict.ContainsKey(constructorInfo)){
                IConstructorInvoker invoker = CreateConstructorInvoker(constructorInfo);
                m_ConstructorInvokerDict.Add(constructorInfo, invoker);
                return invoker;
            }

            return m_ConstructorInvokerDict[constructorInfo];
        }

        public static IMethodInvoker GetMethodInvoker(MethodInfo methodInfo)
        {
            if (!m_MethodInvokerDict.ContainsKey(methodInfo)){
                IMethodInvoker invoker = CreateMethodInvoker(methodInfo);
                m_MethodInvokerDict.Add(methodInfo, invoker);
                return invoker;
            }

            return m_MethodInvokerDict[methodInfo];
        }

        #region Private Methods

        private static IValueSetter CreatePropertySetter(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
                throw new ArgumentException("Please confirm whether propertyInfo is null!");
            if (!propertyInfo.CanWrite)
                throw new ArgumentException("Please confirm whether propertyInfo support write!");

            MethodInfo methodInfo = propertyInfo.GetSetMethod(true);
            Type instanceType = typeof(
        }

        private static IValueGetter CreatePropertyGetter(PropertyInfo propertyInfo)
        {
            throw new NotImplementedException();
        }

        private static IConstructorInvoker CreateConstructorInvoker(ConstructorInfo constructorInfo)
        {
            throw new NotImplementedException();
        }

        private static IMethodInvoker CreateMethodInvoker(MethodInfo methodInfo)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
