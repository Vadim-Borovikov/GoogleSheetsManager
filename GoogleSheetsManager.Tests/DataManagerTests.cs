using System;
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
    public async Task GetTitlesAsync()
    {
        List<string> titles = await DataManager.GetTitlesAsync(_provider, RangeGet);
        Assert.AreEqual(4, titles.Count);
    }

    [TestMethod]
    public async Task GetValuesAsyncTest()
    {
        SheetData<TestInstance> data = await DataManager.GetValuesAsync<TestInstance>(_provider, RangeGet);
        Assert.AreEqual(4, data.Titles.Count);
        Assert.AreEqual(3, data.Instances.Count);

        Assert.AreEqual(TestInstance.Bool, data.Instances[0].Bool);
        Assert.AreEqual(TestInstance.Bool, data.Instances[1].Bool);
        Assert.AreEqual(TestInstance.Bool, data.Instances[2].Bool);

        Assert.AreEqual(TestInstance.Int, data.Instances[0].Int);
        Assert.IsNull(data.Instances[1].Int);
        Assert.AreEqual(TestInstance.Int, data.Instances[2].Int);

        Assert.AreEqual(TestInstance.String, data.Instances[0].String);
        Assert.AreEqual(TestInstance.String, data.Instances[1].String);
        Assert.AreEqual(TestInstance.String, data.Instances[2].String);

        Assert.AreEqual(TestInstance.Uri?.AbsoluteUri, data.Instances[0].Uri?.AbsoluteUri);
        Assert.AreEqual(TestInstance.Uri?.AbsoluteUri, data.Instances[1].Uri?.AbsoluteUri);
        Assert.IsNull(data.Instances[2].Uri);
    }

    [TestMethod]
    public async Task UpdateValuesAsyncTest()
    {
        SheetData<TestInstance> data = await DataManager.GetValuesAsync<TestInstance>(_provider, RangeUpdate);

        // ReSharper disable once NullableWarningSuppressionIsUsed
        data.Instances[0].String = null!;

        await DataManager.UpdateValuesAsync(_provider, RangeGet, data);
        data = await DataManager.GetValuesAsync<TestInstance>(_provider, RangeUpdate);
        Assert.AreEqual(0, data.Instances.Count);

        data.Instances.Add(TestInstance);
        await DataManager.UpdateValuesAsync(_provider, RangeGet, data);
        data = await DataManager.GetValuesAsync<TestInstance>(_provider, RangeUpdate);
        Assert.AreEqual(1, data.Instances.Count);
        Assert.AreEqual(TestInstance.String, data.Instances[0].String);
    }

    [ClassCleanup]
    public static void ClassCleanup() => _provider.Dispose();

    // ReSharper disable once NullableWarningSuppressionIsUsed
    //   _provider initializes in ClassInitialize
    private static SheetsProvider _provider = null!;
    private const string ApplicationName = "GoogleSheetManagerTest";
    private const string RangeGet = "Test!A1:D";
    private const string RangeUpdate = "Test!A1:D2";
    private static readonly TestInstance TestInstance = new()
    {
        Bool = false,
        Int = 32,
        String = "fdgsdfg",
        Uri = new Uri("https://ya.ru"),
    };
}