// See https://aka.ms/new-console-template for more information

using System.Data;
using CNNvs;
using MathWorks.MATLAB.NET.Arrays;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Threading;

class Program
{
    static DateTime lastRunEndTime;

    static void Main()
    {
        // 初始化上次运行结束的时间为当前时间
        lastRunEndTime = DateTime.Now;

        // 创建一个新的线程执行主循环
        Thread loopThread = new Thread(MainLoop);
        loopThread.Start();

        // 按空格键停止循环
        while (Console.ReadKey().Key != ConsoleKey.Spacebar)
        {
            // 等待按下空格键，不做任何操作
        }

        // 设置停止循环的标志为 true
        stopLoop = true;

        // 等待主循环线程结束
        loopThread.Join();

        Console.WriteLine("循环已停止");
    }

    static bool stopLoop = false;

    static void MainLoop()
    {
        while (!stopLoop)
        {
            // 获取当前时间
            DateTime now = DateTime.Now;

            // 计算时间间隔
            TimeSpan interval = now - lastRunEndTime;

            // 判断是否到达指定的间隔时间
            if (interval.TotalMinutes >= 0.5)
            {
                //主体代码，上面是定时器
                #region 读取Excel数据
                /// <summary>
                /// 将excel中的数据导入到DataTable中
                /// </summary>
                /// <param name="fileName">文件路径</param>
                /// <param name="sheetName">excel工作薄sheet的名称</param>
                /// <param name="isFirstRowColumn">第一行是否是DataTable的列名，true是</param>
                /// <returns>返回的DataTable</returns>public 
                static double[,] ExcelToDoubleArray(string fileName, string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            double[,] data = new double[1, 4];//这个定义是data返回结果的类型，现在是1行4列的双精度数组
            FileStream fs;
            IWorkbook workbook = null;
            int cellCount = 0;//列数
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                {
                    workbook = new HSSFWorkbook(fs);
                }
                if (sheetName != null)
                {

                    sheet = workbook.GetSheet(sheetName);//根据给定的sheet名称获取数据
                }
                else
                {
                    //也可以根据sheet编号来获取数据
                    sheet = workbook.GetSheetAt(0);//获取第几个sheet表（此处表示如果没有给定sheet名称，默认是第一个sheet表）  
                }
                if (sheet != null)
                {
                    int rowlast = sheet.LastRowNum;
                    IRow lastRow = sheet.GetRow(rowlast);//读取最后一行数据
                    cellCount = lastRow.LastCellNum; //最后一行最后一个cell的编号 即总的列数
                    if (cellCount != 4)//检测所读取数据数量与预期数量是否一致
                    {
                        Console.Write("读取温度数据与测量数量不符");
                    }
                    //将excel表最后一行的数据添加到datatable的行中
                    double[,] rowData = new double[1, cellCount];
                    for (int j = lastRow.FirstCellNum; j < cellCount; ++j)
                    {
                        NPOI.SS.UserModel.ICell cell = lastRow.GetCell(j);
                        if (lastRow.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                        {
                            rowData[0, j] = cell.NumericCellValue; // 将单元格的数值赋值给数组的对应位置
                        }
                        else
                        {
                            rowData[1, j] = 0.0; // 如果单元格为空或不是数值类型，则将该值置为 0.0 作为默认值
                        }
                    }
                    data = rowData;
                }
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        #endregion
        // 调用上面的方法读取Excel数据并存储在DataTable中
        double[,] dt = ExcelToDoubleArray(@"C:\Users\chengeng\Desktop\用于预测的十五项温度数据.xlsx", "Sheet1", true);//现在存储返回结果的dt就是1行4列的double型数组
                                                                                                          //将1行4列的double型数据表转化成4行1列的double型数组
        int rows = dt.GetLength(0); // 获取原始数组的行数
        int columns = dt.GetLength(1); // 获取原始数组的列数
        double[,] inputData = new double[columns, rows];//创建4行1列的数组作为温度的输入数组
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                inputData[j, i] = dt[i, j]; // 将原始数组中的元素按照转置关系赋值给新数组
            }
        }

        //调用dll以上面处理好的inputDate数据作为输入进行预测
        MWNumericArray wd = new MWNumericArray(inputData);
        CNNvs.myMATLAB predict = new CNNvs.myMATLAB(); //实例化对象
        MWArray e = (MWNumericArray)predict.predictCNN(wd);//调用DLL
                                                           // e的类型由MWARRY转为double
        MWNumericArray arr = (MWNumericArray)e;//强制将类型转换为MWNumericArray
        int numCols = arr.Dimensions[0];
        int numRows = arr.Dimensions[1];

        float[,] arrFloat = (float[,])arr.ToArray(MWArrayComponent.Real);//从MWNumericArray转换为float型变量arrFloat
        double[,] wc = new double[numCols, numRows];//创建double型数组变量wc
        for (int i = 0; i < numCols; i++)
        {
            for (int j = 0; j < numRows; j++)
            {
                wc[i, j] = arrFloat[i, j];
            }
        }

        for (int i = 0; i < numCols; i++)
        {
            for (int j = 0; j < numRows; j++)
            {
                Console.Write(wc[i, j] + " ");//逐个元素打印
            }
            Console.WriteLine();//打印一个回车
        }
          
        
        //
        // 更新上次运行结束的时间为当前时间，上面的是主体函数部分
                lastRunEndTime = now;
            }

            // 休眠一段时间，避免过于频繁检查
            Thread.Sleep(1000); // 这里假设 1000 毫秒（1 秒）为一个适当的间隔
        }
    }
}