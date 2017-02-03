using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
 

namespace ExcelLite.Config
{
    /// <summary>
    /// 基于特性的映射机制实现
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
   public class MappingAttribute:Attribute
    {
        /// <summary>
        ///  映射配置选项
        /// </summary>
        public MappingOption MappingOption{get;private set;}
        /// <summary>
        /// 配置映射键名
        /// </summary>
        public string MappingKey {get;private set;}

        public string PropertyType{get;private set;}

        public string ReferenceStyle{get;private set;}

        public MappingAttribute(MappingOption mappingOption,string mappingKey,string referenceStyle="")
        {
             this.MappingOption = mappingOption;
            this.MappingKey = mappingKey;
            this.ReferenceStyle = referenceStyle; 
        
        }
       

    }
    [Flags]
    public enum MappingOption
    {
        FieldName,
        ColumnName,
        ColumnIndex

    }
}
