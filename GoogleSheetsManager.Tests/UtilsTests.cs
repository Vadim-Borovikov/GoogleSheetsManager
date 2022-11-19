using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using GoogleSheetsManager.Providers;
using Microsoft.Extensions.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GoogleSheetsManager.Tests;

[TestClass]
public class UtilsTests
{
    [ClassInitialize]
    public static void ClassInitialize(TestContext _)
    {
        Config? config = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
                                                   // Create appsettings.json for private settings
                                                   .AddJsonFile("appsettings.json")
                                                   .Build()
                                                   .Get<Config>();

        Assert.IsNotNull(config?.GoogleCredential);
        string googleCredentialJson = JsonSerializer.Serialize(config.GoogleCredential);
        Assert.IsFalse(string.IsNullOrWhiteSpace(googleCredentialJson));

        Assert.IsFalse(string.IsNullOrWhiteSpace(config.GoogleSheetId));
        _provider = new SheetsProvider(googleCredentialJson, ApplicationName, config.GoogleSheetId);
    }

    [TestMethod]
    public async Task LoadTitlesAsyncTest()
    {
        List<string> titles = await Utils.LoadTitlesAsync(_provider, RangeGet);
        Assert.AreEqual(4, titles.Count);
    }

    [ClassCleanup]
    public static void ClassCleanup() => _provider.Dispose();

    // ReSharper disable once NullableWarningSuppressionIsUsed
    //   _provider initializes in ClassInitialize
    private static SheetsProvider _provider = null!;
    private const string ApplicationName = "GoogleSheetManagerTest";
    private const string RangeGet = "Test!A1:D";
}