using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using GoogleSheetsManager.Documents;
using Microsoft.Extensions.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
// ReSharper disable NullableWarningSuppressionIsUsed

namespace GoogleSheetsManager.Tests;

[TestClass]
public class SheetTests
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
        _manager = new DocumentsManager(config);
        Document document = _manager.GetOrAdd(config.GoogleSheetId);
        _sheet = document.GetOrAddSheet(SheeetTitle);
    }

    [TestMethod]
    public async Task LoadTitlesAsyncTest()
    {
        List<string> titles = await _sheet.LoadTitlesAsync(RangeGet);
        Assert.AreEqual(4, titles.Count);
    }

    [TestMethod]
    public async Task LoadAsyncTest()
    {
        SheetData<TestInstance> data = await _sheet.LoadAsync<TestInstance>(RangeGet);
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
        SheetData<TestInstance> data = await _sheet.LoadAsync<TestInstance>(RangeUpdate);

        // ReSharper disable once NullableWarningSuppressionIsUsed
        data.Instances[0].String1 = null!;

        await _sheet.SaveAsync(RangeGet, data);
        data = await _sheet.LoadAsync<TestInstance>(RangeUpdate);
        Assert.AreEqual(0, data.Instances.Count);

        data.Instances.Add(TestInstance);
        await _sheet.SaveAsync(RangeGet, data);
        data = await _sheet.LoadAsync<TestInstance>(RangeUpdate);
        Assert.AreEqual(1, data.Instances.Count);
        Assert.AreEqual(TestInstance.String1, data.Instances[0].String1);
    }

    [ClassCleanup]
    public static void ClassCleanup() => _manager.Dispose();

    private static DocumentsManager _manager = null!;
    private static Sheet _sheet = null!;
    private const string SheeetTitle = "Test";
    private const string RangeGet = "A1:D";
    private const string RangeUpdate = "A1:D2";
    private static readonly TestInstance TestInstance = new()
    {
        Bool = false,
        Int = 32,
        String1 = "fdgsdfg",
        String2 = "dlfnsdkfjgnsd"
    };
}