package com.thinta.readwrite_excel;

public class Box {
	/**
	 * 计量箱条码编号 （资产编号）
	 *
	 */
	public String box_code;
	/**
	 * 表箱类型:
	 *
	 * 单体表箱 合体表箱 虚拟表箱
	 */
	public String box_type;

	/**
	 * 表箱安装地址: 计量箱的安装地址（简要描述关键安装地址信息）
	 */
	public String box_address;

	/**
	 * 表箱行: 计量箱（柜）行
	 */
	public String box_row;

	/**
	 * 表箱列: 计量箱（柜）列
	 */
	public String box_column;

	/**
	 * 表箱材质: 铁、不锈钢、合金、塑料
	 */
	public String box_material;

	/**
	 * 接入点
	 */
	public String box_line;

	/**
	 * 连接设备类型
	 */
	public String box_couplingdev_type ;

	/**
	 * 连接设备名称
	 */
	public String box_couplingdev_name;
}
