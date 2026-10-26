using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Localization;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;
using Sr = TemplateFrame.Excel.Simple.Localization.Sr;

namespace TemplateFrame.Excel.Simple;

public abstract class SimpleExcelTemplateService<TData>
{
    private readonly Lazy<TemplateContract> _contract;

    protected SimpleExcelTemplateService()
    {
        _contract = new Lazy<TemplateContract>(() =>
        {
            var contract = DefineContract() ?? throw new InvalidOperationException(Sr.Get("SimpleExcel.Service.DefineContractNull"));
            var table = SimpleExcelContract.RequireSingleTable(contract);
            if (string.IsNullOrWhiteSpace(table.DataPath) && !DataPathMapper.IsCollectionDataType(typeof(TData)))
            {
                throw new InvalidOperationException(
                    Sr.Get("SimpleExcel.Service.TableNeedsDataPath", contract.Name, table.Key));
            }

            return contract;
        });
    }

    public TemplateContract Contract => _contract.Value;

    protected abstract TemplateContract DefineContract();

    public Stream BuildTemplate(SimpleExcelOptions? options = null, CultureInfo? culture = null, ITemplateLocalizer? localizer = null)
    {
        var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, new FillData(), Contract, options, culture, localizer);
        stream.Position = 0;
        return stream;
    }

    public TemplateValidationResult Validate(Stream template, SimpleExcelOptions? options = null)
        => SimpleExcelContract.Validate(template, Contract, options);

    public Stream Fill(TData data, SimpleExcelOptions? options = null, CultureInfo? culture = null, ITemplateLocalizer? localizer = null)
    {
        var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, MapToData(data), Contract, options, culture, localizer);
        stream.Position = 0;
        return stream;
    }

    public TData Parse(Stream source, SimpleExcelOptions? options = null)
        => MapFromData(SimpleExcelContract.Read(source, Contract, options));

    protected virtual FillData MapToData(TData data)
        => DataPathMapper.ToFillData(data, Contract);

    protected virtual TData MapFromData(FillData data)
        => DataPathMapper.FromFillData<TData>(data, Contract);
}
