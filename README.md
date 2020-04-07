# csv-to-xls
C#/.NET Excel CSV to xls

```bash
  public static DataTable OpenCSV(string filePath)
 {
     Encoding encoding = Encoding.UTF8; //Encoding.ASCII;//
     DataTable dt = new DataTable();
     FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);

     //StreamReader sr = new StreamReader(fs, Encoding.UTF8);
     StreamReader sr = new StreamReader(fs, encoding);
     //string fileContent = sr.ReadToEnd();
     //encoding = sr.CurrentEncoding;
     //记录每次读取的一行记录
     string strLine = "";
     //记录每行记录中的各字段内容
     string[] aryLine = null;
     string[] tableHead = null;
     //标示列数
     int columnCount = 0;
     //标示是否是读取的第一行
     bool IsFirst = true;
     //逐行读取CSV中的数据
     while ((strLine = sr.ReadLine()) != null)
     {
         //strLine = Common.ConvertStringUTF8(strLine, encoding);
         //strLine = Common.ConvertStringUTF8(strLine);

         if (IsFirst == true)
         {
             tableHead = strLine.Split(',');
             IsFirst = false;
             columnCount = tableHead.Length;
             //创建列
             for (int i = 0; i < columnCount; i++)
             {
                 string columnName = tableHead[i];
                 DataColumn dc = new DataColumn(columnName);
                 dt.Columns.Add(dc);
             }
         }
         else
         {
             if (!string.IsNullOrEmpty(strLine))
             {
                 aryLine = strLine.Split(',');
                 if (aryLine.Length != 7)
                 {
                     DataRow dr = dt.NewRow();
                     for (int j = 0; j < columnCount; j++)
                     {
                         dr[j] = aryLine[j].Replace("`", "");
                     }
                     dt.Rows.Add(dr);
                 }
             }
         }
     }
     if (aryLine != null && aryLine.Length > 0)
     {
         if (aryLine.Length != 7)
             dt.DefaultView.Sort = tableHead[0] + " " + "asc";
     }

     sr.Close();
     fs.Close();
     return dt;
 }

public static bool DataTableToExcel(DataTable list)
  {
      string fname2 = "d:\\";
      string fname1 = "table_" + DateTime.Now.ToString("yyyyMMddHHmm");
      string FileName = fname2 + fname1 + ".xls";
      if (!File.Exists(FileName))
      {
          bool result = false;
          IWorkbook workbook = null;
          FileStream fs = null;
          IRow row = null;
          ISheet sheet = null;
          ICell cell = null;
          try
          {
              int s = list.Rows.Count;
              if (list != null && list.Rows.Count > 0)
              {
                  workbook = new HSSFWorkbook();
                  sheet = workbook.CreateSheet("Sheet1");//创建一个名称为Sheet0的表
                  int rowCount = list.Rows.Count;//行数
                  int columnCount = list.Columns.Count;//列数

                  //设置列头
                  row = sheet.CreateRow(0);//excel第一行设为列头
                  for (int c = 0; c < columnCount; c++)
                  {
                      cell = row.CreateCell(c);
                      cell.SetCellValue(list.Columns[c].ColumnName);
                  }

                  //设置每行每列的单元格,
                  for (int i = 0; i < rowCount; i++)
                  {
                      row = sheet.CreateRow(i + 1);
                      for (int j = 0; j < columnCount; j++)
                      {
                          cell = row.CreateCell(j);//excel第二行开始写入数据
                          cell.SetCellValue(list.Rows[i][j].ToString());
                      }
                  }
                  using (fs = File.OpenWrite(FileName))
                  {
                      workbook.Write(fs);//向打开的这个xls文件中写入数据
                      result = true;
                  }
              }
              return result;
          }
          catch (Exception ex)
          {
              if (fs != null)
              {
                  fs.Close();
              }
              return false;
          }
      }
      else
      {
          return false;
      }
  }
-------------------------
//Execute
string filePath = @"D:\123456.csv";
DataTable dt = OpenCSV(filePath);
DataTableToExcel(dt);
```
