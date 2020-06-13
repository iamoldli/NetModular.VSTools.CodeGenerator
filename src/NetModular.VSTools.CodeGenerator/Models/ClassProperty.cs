using System.Collections.Generic;

namespace NetModular.VSTools.CodeGenerator
{
    /// <summary>
    /// 属性
    /// </summary>
    public class ClassProperty
    {
        /// <summary>
        /// 属性类型
        /// </summary>
        public string PropertyType { get; set; }

        /// <summary>
        /// 属性名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 属性中文名称
        /// </summary>
        public string CnName { get; set; }

        /// <summary>
        /// 属性特性
        /// </summary>
        public List<ClassAttribute> ClassAttributes { get; set; }
    }

}
