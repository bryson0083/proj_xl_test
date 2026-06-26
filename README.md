# templexl

> 基於 [openpyxl](https://openpyxl.readthedocs.io/) 的 Excel 模板報表引擎——以 Excel 檔作為模板，將資料渲染成報表。

`templexl` 讓你用熟悉的 Excel 設計報表版面（樣式、表格、公式、圖片），再以 `{{變數}}`
與 `#{{資料表}}` 標籤填入資料，輸出成最終報表。定位為 [xlwings Reports](https://docs.xlwings.org/en/stable/reports/index.html)
的免費替代品，**無需安裝 Excel** 即可使用。

> ⚠️ 開發中（pre-release）。API 在 1.0 之前可能變動。

## 安裝

```bash
pip install templexl
```

需求：Python 3.12+。

## 快速開始

```python
import pandas as pd
from templexl import render

result = render(
    template="template.xlsx",
    output="output.xlsx",
    data={
        "oper_name": "王小明",
        "date_rng_desc": "2025/01/01 - 2025/01/31",
        "report_df": pd.DataFrame({"姓名": ["Alice", "Bob"], "部門": ["技術部", "業務部"]}),
    },
)

print(result.output_path)   # 'output.xlsx'
print(result.warnings)      # 非致命警告（如未解析的標籤）
```

## 模板語法

| 語法 | 說明 |
|------|------|
| `{{變數名稱}}` | 以 `data` 中對應的純量值取代 |
| `#{{資料表名稱}}` | 以 pandas DataFrame 展開為表格（含欄位標題列） |
| `#{{資料表名稱 \| noheader}}` | 展開但略過欄位標題列 |

渲染時會保留模板的樣式、合併儲存格與 Excel 表格（Table）範圍，並平移位於展開
區域下方的公式與圖片。

## API

### `render(template, output, data=None, *, with_report=False) -> RenderResult`

| 參數 | 說明 |
|------|------|
| `template` | 模板檔路徑（`.xlsx` 或 `.xlsm`） |
| `output` | 輸出檔路徑 |
| `data` | 渲染資料字典；鍵為標籤名稱，值為純量或 `pandas.DataFrame` |
| `with_report` | 是否產生除錯用渲染報告（記憶體物件，預設關閉） |

回傳 `RenderResult`：

- `output_path`：已寫出的輸出檔路徑
- `warnings`：非致命警告清單
- `report`：渲染報告（僅 `with_report=True` 時，否則 `None`）；可 `report.write_json(path)` 自行落地

例外皆繼承自 `TemplateError`：`TemplateNotFoundError`、`FileFormatError`、`RenderError`。

```python
from templexl import TemplateError

try:
    render("t.xlsx", "o.xlsx", data={...})
except TemplateError as e:
    ...
```

## 已知限制

- **公式平移**：目前以列向參照調整為主。含數字的函式名（如 `LOG10`）、跨工作表
  參照等情境尚有已知缺陷，詳見測試中標記為「已知缺陷」的案例。
- **圖片**：僅平移既有圖片的列向錨點（保留 rowOff/colOff），不處理欄向位移。
- **大量資料**：渲染耗時隨列數近似線性（100k 列約 15 秒）；採 openpyxl 記憶體模式，
  極大資料量（數十萬列以上）峰值記憶體較高。
- **變數位置**：少數位於展開表格下方的純量標籤可能未被替換，會列入 `warnings`。

## 開發

```bash
uv sync          # 建立虛擬環境並安裝相依（含 dev 群組）
uv run pytest    # 執行黃金檔案回歸測試
```

測試採「黃金檔案」策略：以基準 `.xlsx` 逐格比對渲染輸出（值、樣式、合併、表格
範圍、公式、圖片錨點），確保重構不改變行為。

## 授權

[MIT](LICENSE) © Bryson Xue
