using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using GoogleSheetsManager.Providers;
using Microsoft.Extensions.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace GoogleSheetsManager.Tests;

[TestClass]
public class DataManagerTests
{
    [ClassInitialize]
    public static void ClassInitialize(TestContext _)
    {
        ConfigJson config = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
                                                         // Create appsettings.json for private settings
                                                         .AddJsonFile("appsettings.json")
                                                         .Build()
                                                         .Get<ConfigJson>();

        Assert.IsNotNull(config.GoogleCredential);
        string googleCredentialJson = JsonConvert.SerializeObject(config.GoogleCredential);
        Assert.IsFalse(string.IsNullOrWhiteSpace(googleCredentialJson));

        Assert.IsFalse(string.IsNullOrWhiteSpace(config.GoogleSheetId));
        _provider = new SheetsProvider(googleCredentialJson, ApplicationName, config.GoogleSheetId);
    }

    [TestMethod]
    public async Task UpdateValuesAsyncTest()
    {
        IList<TestInstance> instances = await DataManager.GetValuesAsync(_provider, TestInstance.Load, Range);
        Assert.AreEqual(1, instances.Count);
        Assert.AreEqual(Value, instances[0].Value);

        instances[0].Value = null;
        await DataManager.UpdateValuesAsync(_provider, Range, instances);
        instances = await DataManager.GetValuesAsync(_provider, TestInstance.Load, Range);
        Assert.AreEqual(0, instances.Count);

        TestInstance instance = new() { Value = Value };
        instances.Add(instance);
        await DataManager.UpdateValuesAsync(_provider, Range, instances);
        instances = await DataManager.GetValuesAsync(_provider, TestInstance.Load, Range);
        Assert.AreEqual(1, instances.Count);
        Assert.AreEqual(Value, instances[0].Value);

        _provider.Dispose();
    }

    [ClassCleanup]
    public static void ClassCleanup() => _provider.Dispose();

    // ReSharper disable once NullableWarningSuppressionIsUsed
    //   _provider initializes in ClassInitialize
    private static SheetsProvider _provider = null!;
    private const string ApplicationName = "GoogleSheetManagerTest";
    private const string Range = "Test!A1:2";
    private const string Value = "x";
}