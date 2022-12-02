using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using GoogleSheetsManager.Providers;
using Microsoft.Extensions.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GoogleSheetsManager.Tests;

[TestClass]
public class DataManagerTests
{
    [ClassInitialize]
    public static void ClassInitialize(TestContext _)
    {
        Config? config = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
                                                   // Create appsettings.json for private settings
                                                   .AddJsonFile("appsettings.json")
                                                   .Build()
                                                   .Get<Config>();
        string? json = (config as IConfigGoogleSheets)?.GetCredentialJson();
        Assert.IsFalse(string.IsNullOrWhiteSpace(json));
        Assert.IsFalse(string.IsNullOrWhiteSpace(config?.GoogleSheetId));
        _provider = new SheetsProvider(config, config.GoogleSheetId);
    }

    [TestMethod]
    public async Task LoadTitlesAsyncTest()
    {
        List<string> titles = await Utils.LoadTitlesAsync(_provider, RangeGet);
        Assert.AreEqual(4, titles.Count);
    }

    [TestMethod]
    public async Task LoadAsyncTest()
    {
        SheetData<TestInstance> data = await DataManager<TestInstance>.LoadAsync(_provider, RangeGet);
        Assert.AreEqual(4, data.Titles.Count);
        Assert.AreEqual(3, data.Instances.Count);

        Assert.AreEqual(TestInstance.Bool, data.Instances[0].Bool);
        Assert.AreEqual(TestInstance.Bool, data.Instances[1].Bool);
        Assert.AreEqual(TestInstance.Bool, data.Instances[2].Bool);

        Assert.AreEqual(TestInstance.Int, data.Instances[0].Int);
        Assert.IsNull(data.Instances[1].Int);
        Assert.AreEqual(TestInstance.Int, data.Instances[2].Int);

        Assert.AreEqual(TestInstance.String1, data.Instances[0].String1);
        Assert.AreEqual(TestInstance.String1, data.Instances[1].String1);
        Assert.AreEqual(TestInstance.String1, data.Instances[2].String1);

        Assert.AreEqual(TestInstance.String2, data.Instances[0].String2);
        Assert.AreEqual(TestInstance.String2, data.Instances[1].String2);
        Assert.IsNull(data.Instances[2].String2);
    }

    [TestMethod]
    public async Task SaveAsyncTest()
    {
        SheetData<TestInstance> data = await DataManager<TestInstance>.LoadAsync(_provider, RangeUpdate);

        // ReSharper disable once NullableWarningSuppressionIsUsed
        data.Instances[0].String1 = null!;

        await DataManager<TestInstance>.SaveAsync(_provider, RangeGet, data);
        data = await DataManager<TestInstance>.LoadAsync(_provider, RangeUpdate);
        Assert.AreEqual(0, data.Instances.Count);

        data.Instances.Add(TestInstance);
        await DataManager<TestInstance>.SaveAsync(_provider, RangeGet, data);
        data = await DataManager<TestInstance>.LoadAsync(_provider, RangeUpdate);
        Assert.AreEqual(1, data.Instances.Count);
        Assert.AreEqual(TestInstance.String1, data.Instances[0].String1);
    }

    [ClassCleanup]
    public static void ClassCleanup() => _provider.Dispose();

    // ReSharper disable once NullableWarningSuppressionIsUsed
    //   _provider initializes in ClassInitialize
    private static SheetsProvider _provider = null!;
    private const string RangeGet = "Test!A1:D";
    private const string RangeUpdate = "Test!A1:D2";
    private static readonly TestInstance TestInstance = new()
    {
        Bool = false,
        Int = 32,
        String1 = "fdgsdfg",
        String2 = "dlfnsdkfjgnsd"
    };
}