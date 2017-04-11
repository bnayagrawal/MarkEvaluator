using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace markevaluator
{
    interface IParse
    {
        void ValidateWorksheet(string fileName);
        void PushToDatabase(string fileName);
    }
}
