using System;
using Microsoft.ML;
using Microsoft.ML.Data;

namespace myapp
{
    internal class Program
    {
        // Класс для входных данных
        public class InputData
        {
            [LoadColumn(0)] public string Column1 { get; set; }
            [LoadColumn(1)] public string Column2 { get; set; }
            [LoadColumn(2)] public string Label { get; set; }
        }

        // Класс для предсказанных значений
        public class OutputData
        {
            [ColumnName("PredictedLabel")] public string MatchCategory { get; set; }
        }

        static void Main()//Samsung Galaxy A52
        {
            string modelPath = "model.zip"; // Файл обученной модели
            var mlContext = new MLContext();

            // Загружаем модель
            ITransformer model = mlContext.Model.Load(modelPath, out _);
            var predictor = mlContext.Model.CreatePredictionEngine<InputData, OutputData>(model);

            Console.WriteLine("Введите первую строку:");
            string input1 = Console.ReadLine();

            Console.WriteLine("Введите вторую строку:");
            string input2 = Console.ReadLine();

            // Делаем предсказание
            var inputData = new InputData { Column1 = input1, Column2 = input2 };
            var result = predictor.Predict(inputData);

            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine($"Результат: {result.MatchCategory}");
        }
    }
}
