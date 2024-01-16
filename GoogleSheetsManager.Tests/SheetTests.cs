using System.Collections.Generic;
using System.IO;
using System.Text.Json;
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
        Assert.IsNotNull(config);
        string credentialJson = JsonSerializer.Serialize(config!.Credential);
        Assert.IsFalse(string.IsNullOrWhiteSpace(credentialJson));
        Assert.IsFalse(string.IsNullOrWhiteSpace(config.GoogleSheetId));
        _id = config.GoogleSheetId;
        _manager = new Manager(config);
        Document document = _manager.GetOrAdd(_id);
        _sheet = document.GetOrAddSheet(SheeetTitle);
    }

    [TestMethod]
    public async Task DownloadAsyncTest()
    {
        string path = Path.GetTempFileName();

        await _manager.DownloadAsync(_id, PdfMime, path);
        Assert.IsTrue(File.Exists(path));
        byte[] bytes = await File.ReadAllBytesAsync(path);
        Assert.AreEqual(PdfSize, bytes.Length);
        File.Delete(path);
    }

    [TestMethod]
    public async Task LoadTitlesAsyncTest()
    {
        await _sheet.LoadTitlesAsync(RangeGet);
        Assert.AreEqual(4, _sheet.Titles.Count);
    }

    [TestMethod]
    public async Task LoadAsyncTest()
    {
        List<TestInstance> data = await _sheet.LoadAsync<TestInstance>(RangeGet);
        Assert.AreEqual(4, _sheet.Titles.Count);
        Assert.AreEqual(2, data.Count);

        Assert.AreEqual(TestInstance.Bool, data[0].Bool);
        Assert.AreEqual(TestInstance.Bool, data[1].Bool);

        Assert.AreEqual(TestInstance.Int, data[0].Int);
        Assert.IsNull(data[1].Int);

        Assert.AreEqual(TestInstance.String1, data[0].String1);
        Assert.AreEqual(TestInstance.String1, data[1].String1);

        Assert.AreEqual(TestInstance.String2, data[0].String2);
        Assert.IsNull(data[1].String2);

        data = await _sheet.LoadAsync<TestInstance>(RangeGetEmpty);
        Assert.AreEqual(4, _sheet.Titles.Count);
        Assert.AreEqual(0, data.Count);
    }

    [TestMethod]
    public async Task SaveAsyncTest()
    {
        List<TestInstance> data = await _sheet.LoadAsync<TestInstance>(RangeUpdate);

        // ReSharper disable once NullableWarningSuppressionIsUsed
        data[0].String1 = null!;

        await _sheet.SaveAsync(RangeGet, data);
        data = await _sheet.LoadAsync<TestInstance>(RangeUpdate);
        Assert.AreEqual(0, data.Count);

        data.Add(TestInstance);
        await _sheet.SaveAsync(RangeGet, data);
        data = await _sheet.LoadAsync<TestInstance>(RangeUpdate);
        Assert.AreEqual(1, data.Count);
        Assert.AreEqual(TestInstance.String1, data[0].String1);
    }

    [ClassCleanup]
    public static void ClassCleanup() => _manager.Dispose();

    private static string _id = null!;
    private static Manager _manager = null!;
    private static Sheet _sheet = null!;
    private const string SheeetTitle = "Test";
    private const string RangeGet = "A1:D";
    private const string RangeGetEmpty = "A1:D1";
    private const string RangeUpdate = "A1:D2";
    private const string PdfMime = "application/pdf";
    private const int PdfSize = 791528;
    private static readonly TestInstance TestInstance = new()
    {
        Bool = false,
        Int = 32,
        String1 = "fdgsdfg",
        String2 = "dlfnsdkfjgnsd"
    };
}