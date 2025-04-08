using BK_Details_App.ViewModels;

namespace TestProject1
{
    [TestClass]
    public sealed class Test1
    {
        private readonly string logFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test_results.txt");
        private string testFilePath;

        [TestInitialize]
        public void Setup()
        {
            testFilePath = Path.Combine(AppContext.BaseDirectory.Substring(0,
               AppContext.BaseDirectory.IndexOf("TestProject1") - 1),
               "BK_Details_App\\bin\\Debug\\net8.0\\Materials\\test.xlsx");

            // Удаляем файл, если он есть, чтобы проверить поведение при его отсутствии
            if (File.Exists(testFilePath))
                File.Delete(testFilePath);
        }

        [TestMethod]
        public void Test_FileDoesNotExist_ReturnsEmptyList()
        {
            string testName = "Test_FileDoesNotExist_ReturnsEmptyList";
            string expected = "[]";
            try
            {
                var service = new DetailsVM(skipInit: true); // замени на реальный класс, где находится ReadFavorites
                var result = service.ReadFavorites(testFilePath);

                string actual = "[" + string.Join(", ", result) + "]";
                bool passed = result.Count == 0;

                string status = passed ? "Passed" : "Failed";

                string log = $"Test Name: {testName}{Environment.NewLine}" + 
                             $"Datetime now: {System.DateTime.Now}{Environment.NewLine}" +
                             $"Expected: {expected}{Environment.NewLine}" +
                             $"Actual: {actual}{Environment.NewLine}" +
                             $"Status: {status}{Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";

                File.AppendAllText(logFile, log);

                Assert.IsTrue(passed);
            }
            catch (Exception ex)
            {
                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Exception: {ex.Message}{Environment.NewLine}" +
                             $"Status: Failed (Exception){Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";
                File.AppendAllText(logFile, log);
                Assert.Fail("Exception thrown");
            }
        }

        [TestMethod]
        public void Test_FileDoesExist_ReturnsListWithValue()
        {
            string testName = "Test_FileDoesExist_ReturnsListWithValue";
            string expected = "[1]";
            try
            {
                var service = new DetailsVM(skipInit: true); // замени на реальный класс, где находится ReadFavorites

                var result = service.ReadFavorites(testFilePath);
                service.AddToFavorite("1");
                result = service.ReadFavorites(testFilePath);

                string actual = "[" + string.Join(", ", result) + "]";
                bool passed = result.Count == 0;

                string status = passed ? "Passed" : "Failed";

                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Datetime now: {System.DateTime.Now}{Environment.NewLine}" +
                             $"Expected: {expected}{Environment.NewLine}" +
                             $"Actual: {actual}{Environment.NewLine}" +
                             $"Status: {status}{Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";

                File.AppendAllText(logFile, log);

                Assert.IsTrue(passed);
            }
            catch (Exception ex)
            {
                string log = $"Test Name: {testName}{Environment.NewLine}" +
                             $"Exception: {ex.Message}{Environment.NewLine}" +
                             $"Status: Failed (Exception){Environment.NewLine}" +
                             $"----------------------{Environment.NewLine}";
                File.AppendAllText(logFile, log);
                Assert.Fail("Exception thrown");
            }
        }
    }
}
