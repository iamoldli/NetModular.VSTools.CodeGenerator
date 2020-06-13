using System.Collections.Generic;

namespace NetModular.VSTools.CodeGenerator
{
    /// <summary>
    /// 函数
    /// </summary>
    public class ClassFunction
    {
        public string Access { get; set; }

        /// <summary>
        /// 函数类型
        /// </summary>
        public string FunctionType { get; set; }

        /// <summary>
        /// 函数名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 函数参数
        /// </summary>
        public List<ClassParameter> ClassParameters { get; set; }

        /// <summary>
        /// 函数参数
        /// </summary>
        public string ClassParametersCast { get; set; }

    }


}
