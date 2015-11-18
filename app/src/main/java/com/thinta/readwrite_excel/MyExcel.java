package com.thinta.readwrite_excel;

import java.io.File;
import java.util.List;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;


public class MyExcel {

	private WritableWorkbook book;

	String Path;
	// 第一页
	private WritableSheet sheet0;
	// 第二页
	private WritableSheet sheet1;

	public MyExcel(String Path) {
		this.Path = Path;
	}

	/**
	 * 打开Excel，不存在， 创建Excel
	 *
	 * @throws Exception
	 */
	public void open() throws Exception {
		book = Workbook.createWorkbook(new File(Path));
		writeHeadExcel();
	}

	/**
	 * 关闭Excel
	 *
	 * @throws Exception
	 */
	public void close() throws Exception {
		// 从内存中写入文件中
		book.write();
		book.close();
	}

	/**
	 * 入表头格式
	 *
	 * @param
	 * @throws Exception
	 * @throws
	 */
	private void writeHeadExcel() throws Exception {

		// --------------------------------第一页-------------------
		// 生成名为“第一页”的工作表,参数0表示这是第一页

		sheet0 = book.createSheet("4.计量箱（柜）_外业", 0);
		// 第一页列宽
		sheet0.setColumnView(0, 6);
		sheet0.setColumnView(1, 25);
		sheet0.setColumnView(2, 15);
		sheet0.setColumnView(3, 40);
		sheet0.setColumnView(4, 8);
		sheet0.setColumnView(5, 8);
		sheet0.setColumnView(6, 15);
		sheet0.setColumnView(7, 15);
		sheet0.setColumnView(8, 15);
		sheet0.setColumnView(9, 15);

		sheet0.mergeCells(0, 0, 7, 0); // 合并第一行
		// 在Label对象的构造函数中,元格位置是第一列第一行(0,0)以及单元格内容为test
		Label labe_0_0_0 = new Label(0, 0, "计量箱（柜)", setBoldfont(20, true, Colour.BLACK, Colour.WHITE));
		Label labe_0_0_1 = new Label(0, 1, "序号", setBoldfont(10, true, Colour.BLACK, Colour.GRAY_25));
		Label labe_0_1_1 = new Label(1, 1, "计量箱条码编号", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_0_2_1 = new Label(2, 1, "表箱类型", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_0_3_1 = new Label(3, 1, "安装地址", setBoldfont(10, true, Colour.BLACK, Colour.GRAY_25));
		Label labe_0_4_1 = new Label(4, 1, "行", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_0_5_1 = new Label(5, 1, "列", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_0_6_1 = new Label(6, 1, "材质", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_0_7_1 = new Label(7, 1, "接入点", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_0_8_1 = new Label(8, 1, "接入点", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_0_9_1 = new Label(9, 1, "接入点", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));

		// 将定义好的单元格添加到工作表中
		sheet0.addCell(labe_0_0_0);
		sheet0.addCell(labe_0_0_1);
		sheet0.addCell(labe_0_1_1);
		sheet0.addCell(labe_0_2_1);
		sheet0.addCell(labe_0_3_1);
		sheet0.addCell(labe_0_4_1);
		sheet0.addCell(labe_0_5_1);
		sheet0.addCell(labe_0_6_1);
		sheet0.addCell(labe_0_7_1);
		sheet0.addCell(labe_0_8_1);
		sheet0.addCell(labe_0_9_1);

		// --------------------------------第二页-------------------

		sheet1 = book.createSheet("4-1.计量箱与电能表的关系_外业", 1);
		// 第二页列宽
		sheet1.setColumnView(0, 6);
		sheet1.setColumnView(1, 25);
		sheet1.setColumnView(2, 25);
		sheet1.setColumnView(3, 15);
		sheet1.setColumnView(4, 15);

		sheet1.mergeCells(0, 0, 4, 0); // 合并第一行
		// 在Label对象的构造函数中,元格位置是第一列第一行(0,0)以及单元格内容为test
		Label labe_1_0_0 = new Label(0, 0, "计量箱与电能表的关系", setBoldfont(20, true, Colour.BLACK, Colour.WHITE));
		Label labe_1_0_1 = new Label(0, 1, "序号", setBoldfont(10, true, Colour.BLACK, Colour.GRAY_25));
		Label labe_1_1_1 = new Label(1, 1, "计量箱条码编号", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_1_2_1 = new Label(2, 1, "电表条码编号", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_1_3_1 = new Label(3, 1, "电表表位（行）", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));
		Label labe_1_4_1 = new Label(4, 1, "电表表位（列）", setBoldfont(10, true, Colour.RED, Colour.GRAY_25));

		// 将定义好的单元格添加到工作表中
		sheet1.addCell(labe_1_0_0);
		sheet1.addCell(labe_1_0_1);
		sheet1.addCell(labe_1_1_1);
		sheet1.addCell(labe_1_2_1);
		sheet1.addCell(labe_1_3_1);
		sheet1.addCell(labe_1_4_1);
		// book.write();
	}

	/**
	 * 写入数据的样式
	 *
	 * @param sp
	 *            字体大小
	 * @param bl
	 *            是否字体加粗（true 加粗）
	 * @param font_Colour
	 *            字体颜色
	 * @param background_Colour
	 *            背景颜色
	 * @return
	 * @throws Exception
	 */
	private WritableCellFormat setBoldfont(int sp, boolean bl, Colour font_Colour, Colour background_Colour) throws Exception {
		WritableFont wfc;
		if (bl) {
			// 字体 大小 字体样式(加粗) 是否是斜体 是否有下划线 字体颜色
			wfc = new WritableFont(WritableFont.createFont("宋体"), sp, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, font_Colour);
		} else {
			wfc = new WritableFont(WritableFont.createFont("宋体"), sp, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, font_Colour);
		}
		WritableCellFormat wcfFC = new WritableCellFormat(wfc);
		wcfFC.setBackground(background_Colour);// 背景颜色
		wcfFC.setAlignment(Alignment.CENTRE);// 居中
		wcfFC.setBorder(Border.ALL, BorderLineStyle.MEDIUM);// 边线框
		wcfFC.setWrap(true);// 自动换行
		return wcfFC;
	}

	/**
	 * 写数据
	 *
	 * @param box_list
	 * @param box_Ammeters
	 */
	public void WriteData(List<Box> box_list, List<Box_Ammeter> box_Ammeters) {
//		if (box_list.size() == 0 || box_Ammeters.size() == 0) {
//			return;
//		}

		// 起始第3行 行号2
		try {

			for (int i = 0; i < 1000; i++) {
				Label labe0 = new Label(0, i + 2, i + "");
				Label labe1 = new Label(1, i + 2, "a");
				Label labe2 = new Label(2, i + 2, "a");
				Label labe3 = new Label(3, i + 2, "a");
				Label labe4 = new Label(4, i + 2, "a");
				Label labe5 = new Label(5, i + 2, "a");
				Label labe6 = new Label(6, i + 2, "a");
				Label labe7 = new Label(7, i + 2, "a");
				Label labe8 = new Label(8, i + 2, "a");
				Label labe9 = new Label(9, i + 2, "a");

				sheet0.addCell(labe0);
				sheet0.addCell(labe1);
				sheet0.addCell(labe2);
				sheet0.addCell(labe3);
				sheet0.addCell(labe4);
				sheet0.addCell(labe5);
				sheet0.addCell(labe6);
				sheet0.addCell(labe7);
				sheet0.addCell(labe8);
				sheet0.addCell(labe9);
				// book.write();
			}

			// 写表2
			for (int i = 0; i < 1000; i++) {
				Label labe0 = new Label(0, i + 2, i + "");
				Label labe1 = new Label(1, i + 2, "a");
				Label labe2 = new Label(2, i + 2, "a");
				Label labe3 = new Label(3, i + 2, "a");
				Label labe4 = new Label(4, i + 2, "a");

				// 将定义好的单元格添加到工作表中
				sheet1.addCell(labe0);
				sheet1.addCell(labe1);
				sheet1.addCell(labe2);
				sheet1.addCell(labe3);
				sheet1.addCell(labe4);
			}

		} catch (Exception e) {
			// TODO: handle exception
		}

	}

}
