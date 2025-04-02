using System.Data;
using Microsoft.ML;
using Microsoft.ML.Data;

namespace trener
{
    internal class Program
    {
        public class InputData
        {
            [LoadColumn(0)] public string Name { get; set; }
            [LoadColumn(1)] public string Gender { get; set; }
            [LoadColumn(2)] public string Count { get; set; }
            [LoadColumn(3)] public string Probability { get; set; }
            
        }

        public class OutputData
        {
            [ColumnName("PredictedLabel")] public string MatchCategory { get; set; }
        }

        public static void TrainAndSaveModel(string dataPath, string modelPath)
        {
            var mlContext = new MLContext();

            // Загружаем данные из CSV (файл должен иметь заголовки)
            var data = mlContext.Data.LoadFromTextFile<InputData>(path: dataPath, separatorChar: ';', hasHeader: true);
            var split = mlContext.Data.TrainTestSplit(data, testFraction: 0.2);
            var trainData = split.TrainSet;
            var testData = split.TestSet;

            // Определяем пайплайн обработки данных и обучения модели
            var pipeline = mlContext.Transforms.Conversion
                .ConvertType("Count", outputKind: DataKind.Single)
                .Append(mlContext.Transforms.Conversion.ConvertType("Probability", outputKind: DataKind.Single))
                .Append(mlContext.Transforms.Conversion.MapValueToKey("Gender"))
                .Append(mlContext.Transforms.Text.FeaturizeText("FeaturesName", "Name"))
                .Append(mlContext.Transforms.Concatenate(
                    "Features", "FeaturesName", "Count", "Probability"))
                .Append(mlContext.MulticlassClassification.Trainers.SdcaMaximumEntropy("Gender", "Features"))
                .Append(mlContext.Transforms.Conversion.MapKeyToValue("PredictedLabel"));

            Console.WriteLine("Обучение модели...");
            var model = pipeline.Fit(data);

            var predictions = model.Transform(testData);
            var metrics = mlContext.MulticlassClassification.Evaluate(predictions);
            Console.WriteLine($"Accuracy: {metrics.MacroAccuracy:P2}");

            // Сохраняем модель
            mlContext.Model.Save(model, data.Schema, modelPath);
            Console.WriteLine($"Модель сохранена в {modelPath}");
        }

        public static void Main()
        {
            string dataPath = "name_gender_dataset1.csv";

            // Укажи путь к CSV-файлу
            string modelPath = "model.zip";

            // Имя файла модели
            TrainAndSaveModel(dataPath, modelPath);
        }
    }
}
