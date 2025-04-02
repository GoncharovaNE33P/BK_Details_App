using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ML;
using Microsoft.ML.Data;

namespace BK_Details_App.Models
{
    public class a
    {
        public class InputData
        {
            [LoadColumn(0)] public string Column1 { get; set; }
            [LoadColumn(1)] public string Column2 { get; set; }
        }

        public class OutputData
        {
            [ColumnName("PredictedLabel")] public string MatchCategory { get; set; }
        }

        public static void TrainAndSaveModel(string dataPath, string modelPath)
        {
            var mlContext = new MLContext();        
            
            // Загружаем данные из CSV (файл должен иметь заголовки)
            IDataView data = mlContext.Data.LoadFromTextFile<InputData>(dataPath, separatorChar: ',', hasHeader: true);        
            
            // Определяем пайплайн обработки данных и обучения модели
            var pipeline = mlContext.Transforms.Conversion.MapValueToKey("Label").Append(mlContext.Transforms.Text.FeaturizeText("Features", "Column1"))            
                .Append(mlContext.Transforms.Text.FeaturizeText("Features", "Column2"))            
                .Append(mlContext.MulticlassClassification.Trainers.SdcaMaximumEntropy("Label", "Features"))            
                .Append(mlContext.Transforms.Conversion.MapKeyToValue("PredictedLabel"));        
            
            Console.WriteLine("Обучение модели...");        
            var model = pipeline.Fit(data);        
            
            // Сохраняем модель
            mlContext.Model.Save(model, data.Schema, modelPath);        
            Console.WriteLine($"Модель сохранена в {modelPath}");    
        }    
        
        //public static void Main()    
        //{        
        //    string dataPath = "data.csv";   
            
        //    // Укажи путь к CSV-файлу
        //    string modelPath = "model.zip"; 
            
        //    // Имя файла модели
        //    TrainAndSaveModel(dataPath, modelPath);   
        //}
    }
    
    
    
}


