
namespace markevaluator
{
    interface IParse
    {
        void ValidateWorksheet(string fileName);
        void PushToDatabase(string fileName);
    }
}
