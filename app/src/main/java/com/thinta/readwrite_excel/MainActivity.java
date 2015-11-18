package com.thinta.readwrite_excel;

import java.io.File;

import android.os.Bundle;
import android.view.View;
import android.widget.Toast;
import android.app.Activity;

public class MainActivity extends Activity {

	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
	}

	public void Btn_createExcel(View view) {

		try {
			MyExcel myExcel = new MyExcel("/mnt/sdcard/test.xls");
			myExcel.open();
			myExcel.WriteData(null, null);
			myExcel.close();
		} catch (Exception e) {
			Toast.makeText(this, "Exception:"+e.toString(), 0).show();
		}


		File file=new File("/mnt/sdcard/test.xls");
		if (file.exists()) {
			Toast.makeText(this, "test.xls 存在", 0).show();
		}
		else {
			Toast.makeText(this, "test.xls 不存在", 0).show();
		}
	}
}
