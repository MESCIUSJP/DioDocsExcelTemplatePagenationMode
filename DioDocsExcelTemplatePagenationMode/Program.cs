// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;
using System.Data;

Console.WriteLine("DioDocsで明細行数を固定したExcel帳票を作成する");

// 新規ワークブックの作成
var workbook = new Workbook();

// 帳票テンプレートファイルを読み込む
workbook.Open("Template_Pagination.xlsx");

#region データの初期化
DataTable data = new DataTable();
data.Columns.Add("Customer", typeof(string));
data.Columns.Add("CustomerCode", typeof(string));
data.Columns.Add("Product", typeof(string));
data.Columns.Add("Quantity", typeof(int));
data.Columns.Add("Price", typeof(int));

data.Rows.Add("情報システム(株)", "C-001", "カーボン用紙A4", 1, 500);
data.Rows.Add("情報システム(株)", "C-001", "カーボン用紙A3", 2, 510);
data.Rows.Add("情報システム(株)", "C-001", "カーボン用紙A2", 3, 520);
data.Rows.Add("情報システム(株)", "C-001", "給与封筒", 1500, 450);
data.Rows.Add("情報システム(株)", "C-001", "表示用シール（赤）", 10, 250);
data.Rows.Add("情報システム(株)", "C-001", "表示用シール（青）", 5, 250);
data.Rows.Add("情報システム(株)", "C-001", "表示用シール（黄）", 5, 250);
data.Rows.Add("情報システム(株)", "C-001", "ビデオラベル（背見出し）", 2, 500);
data.Rows.Add("情報システム(株)", "C-001", "ビデオラベル（正面用）", 2, 500);
data.Rows.Add("情報システム(株)", "C-001", "プリンタ用トナー", 10, 9000);
data.Rows.Add("情報システム(株)", "C-001", "住所ラベル", 15000, 500);
data.Rows.Add("情報システム(株)", "C-001", "ワープロリボン（黒）", 10, 1000);
data.Rows.Add("情報システム(株)", "C-001", "ワープロリボン（赤）", 15, 1000);
data.Rows.Add("情報システム(株)", "C-001", "ワープロリボン（青）", 20, 1000);
data.Rows.Add("情報システム(株)", "C-001", "A4ファイル", 50, 90);
data.Rows.Add("情報システム(株)", "C-001", "B4ファイル", 30, 90);
data.Rows.Add("情報システム(株)", "C-001", "消しゴム", 20, 50);
data.Rows.Add("旭上株式会社", "C-002", "ボールペン (赤)", 50, 100);
data.Rows.Add("旭上株式会社", "C-002", "ボールペン (青)", 75, 100);
data.Rows.Add("旭上株式会社", "C-002", "ボールペン (緑)", 100, 100);
data.Rows.Add("旭上株式会社", "C-002", "付箋紙（小）", 20, 120);
data.Rows.Add("旭上株式会社", "C-002", "付箋紙（中）", 10, 150);
data.Rows.Add("旭上株式会社", "C-002", "付箋紙（大）", 15, 200);
data.Rows.Add("旭上株式会社", "C-002", "A4コピー用紙", 50, 300);
data.Rows.Add("旭上株式会社", "C-002", "B4コピー用紙", 20, 500);
data.Rows.Add("旭上株式会社", "C-002", "A4ファイル", 5, 90);
data.Rows.Add("旭上株式会社", "C-002", "B4ファイル", 5, 90);
data.Rows.Add("旭上株式会社", "C-002", "クリアケース", 10, 200);
data.Rows.Add("旭上株式会社", "C-002", "クリップ", 30, 50);
data.Rows.Add("旭上株式会社", "C-002", "インク (黒)", 1, 800);
data.Rows.Add("旭上株式会社", "C-002", "インク (赤)", 2, 800);
data.Rows.Add("旭上株式会社", "C-002", "インク (緑)", 3, 800);
#endregion

// 改ページモードをtrueに設定
workbook.Names.Add("TemplateOptions.PaginationMode", "true");

// データソースを追加
workbook.AddDataSource("ds", data);

// テンプレート処理を呼び出し
workbook.ProcessTemplate();

// Excelファイルに保存
workbook.Save("Result.xlsx");

