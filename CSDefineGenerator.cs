using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
namespace excel2json
{
	/// <summary>
	/// 根据表头，生成C#类定义数据结构
	/// 表头使用三行定义：字段名称、字段类型、注释
	/// </summary>
	class CSDefineGenerator
	{
		struct FieldDef
		{
			public string name;
			public string type;
			public string comment;
		}

		string mCode;

		public string code {
			get {
				return this.mCode;
			}
		}

		public CSDefineGenerator (string excelName, ExcelLoader excel, string excludePrefix)
		{
			//-- 创建代码字符串
			StringBuilder sb = new StringBuilder();
			sb.AppendLine("// Auto Generated Code By excel2json");
			sb.AppendLine("// https://neil3d.gitee.io/coding/excel2json.html");
			sb.AppendLine("// 1. 每个 Sheet 形成一个 Struct 定义, Sheet 的名称作为 Struct 的名称");
			sb.AppendLine("// 2. 表格约定：第一行是变量名称，第二行是变量类型");
			sb.AppendLine();
			sb.AppendFormat("// Generate From {0}.xlsx", excelName);
			sb.AppendLine();
			sb.AppendLine("using System;");
			sb.AppendLine("using System.Collections.Generic;");
			sb.AppendLine("using Newtonsoft.Json;");

			for (int i = 0; i < excel.Sheets.Count; i++) {
				DataTable sheet = excel.Sheets[i];
				sb.Append(_exportSheet(sheet, excludePrefix));
			}
			sb.AppendLine("// End of Auto Generated Code");

			// 获取最后一行的起始位置和长度
			int startIndex = sb.ToString().LastIndexOf(Environment.NewLine, StringComparison.Ordinal);
			// 删除最后一行
			sb.Remove(startIndex, sb.Length - startIndex);

			mCode = sb.ToString();
		}

		private string _exportSheet (DataTable sheet, string excludePrefix)
		{
			if (sheet.Columns.Count < 0 || sheet.Rows.Count < 2)
				return "";

			string sheetName = sheet.TableName;
			if (excludePrefix.Length > 0 && sheetName.StartsWith(excludePrefix))
				return "";

			// get field list
			List<FieldDef> fieldList = new List<FieldDef>();
			DataRow typeRow = sheet.Rows[0];
			DataRow commentRow = sheet.Rows[1];

			foreach (DataColumn column in sheet.Columns) {
				// 过滤掉包含指定前缀的列
				string columnName = column.ToString();
				if (excludePrefix.Length > 0 && columnName.StartsWith(excludePrefix) || columnName.StartsWith("#"))
					continue;

				FieldDef field;
				field.name = column.ToString();
				field.type = typeRow[column].ToString();
				field.comment = commentRow[column].ToString();

				fieldList.Add(field);
			}

			// export as string
			StringBuilder sb = new StringBuilder();
			sb.AppendLine("namespace Data\r\n{");
			sb.AppendFormat("\tpublic class {0}\r\n\t{{\n", sheet.TableName);
			sb.AppendFormat("\t\tpublic static Type FormatType() {{ return typeof(Dictionary<uint, {0}>); }}\n\n", sheet.TableName);

			foreach (var field in fieldList) {
				bool isList = field.type.StartsWith("List");
				if (isList) {
					string fieldName = field.name + "List";
					sb.AppendLine("\t\t[JsonIgnore]");
					sb.AppendFormat("\t\tprivate {0} {1}; // {2}\n", field.type, fieldName, field.comment);
					sb.AppendFormat("\t\tpublic string {0} {{ get {{ return JsonConvert.SerializeObject({1}); }} set {{ {2} = JsonConvert.DeserializeObject<{3}>(value); }} }}", field.name, fieldName, fieldName, field.type);
				} else {
					sb.AppendFormat("\t\tpublic {0} {1}; // {2}", field.type, field.name, field.comment);
				}
				sb.AppendLine();
			}

			// 获取最后一行的起始位置和长度
			int startIndex = sb.ToString().LastIndexOf(Environment.NewLine, StringComparison.Ordinal);
			// 删除最后一行
			sb.Remove(startIndex, sb.Length - startIndex);

			sb.Append("\r\t}");
			sb.Append("\n}");
			sb.AppendLine();
			return sb.ToString();
		}

		public void SaveToFile (string filePath, Encoding encoding)
		{
			//-- 保存文件
			using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write)) {
				using (TextWriter writer = new StreamWriter(file, encoding))
					writer.Write(mCode);
			}
		}
	}
}