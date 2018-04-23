using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RichEditAPISample
{
    #region RichEditExampleEvaluatorByTimer
    public class RichEditExampleEvaluatorByTimer : ExampleEvaluatorByTimer
    {
        public RichEditExampleEvaluatorByTimer()
            : base()
        {
        }

        protected override ExampleCodeEvaluator GetExampleCodeEvaluator(ExampleLanguage language)
        {
            if (language == ExampleLanguage.VB)
                return new RichEditVbExampleCodeEvaluator();
            return new RichEditCSExampleCodeEvaluator();
        }
    }
    #endregion
}
