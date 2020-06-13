using System.Collections.Generic;

namespace NetModular.VSTools.CodeGenerator
{
    public class ClassModel
    {
        public string Namespace { get; set; }

        public string Name { get; set; }

        public string CnName { get; set; }

        public string Description { get; set; }

        public string DirName { get; set; }

        public List<ClassFunction> ClassFunctions { get; set; }

        public List<ClassProperty> ClassPropertys { get; set; }
    }
}
