package com.thinta.readwrite_excel;

public class Box_Ammeter {
	/**
	 * 计量箱条码编号（资产编号）
	 */
	public String box_code;

	/**
	 * 电表条码编号: 表箱内所有电能表的条形码编号
	 */
	public String ammeter_code;

	/**
	 * 电表表位（行）: 自上而下（单体表箱行为1）
	 */
	public String ammeter_seat_row;

	/**
	 * 电表表位（列）: 从左向右（单体表箱列为1）
	 */
	public String ammeter_seat_column;
}
