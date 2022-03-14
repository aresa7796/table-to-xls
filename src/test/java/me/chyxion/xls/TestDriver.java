package me.chyxion.xls;

import me.chyxion.xls.model.SheetDTO;
import org.junit.Test;
import java.util.Scanner;
import java.io.FileOutputStream;

/**
 * @version 0.0.1
 * @since 0.0.1
 * @author Shaun Chyxion <br>
 * chyxion@163.com <br>
 * Oct 24, 2014 2:07:51 PM
 */
public class TestDriver {

	@Test
	public void run() throws Exception {
		StringBuilder html = new StringBuilder();
		Scanner s = new Scanner(getClass().getResourceAsStream("/sample.html"), "utf-8");
		while (s.hasNext()) {
			html.append(s.nextLine());
		}
		s.close();

		StringBuilder html1 = new StringBuilder();
		Scanner s1 = new Scanner(getClass().getResourceAsStream("/sample1.html"), "utf-8");
		while (s1.hasNext()) {
			html1.append(s1.nextLine());
		}
		s1.close();
		FileOutputStream fout = new FileOutputStream("target/data.xls");
		SheetDTO[] sheets = new SheetDTO[2];
		sheets[0] = new SheetDTO();
		sheets[0].setSheetName("0000000");
		sheets[0].setHtml(html.toString());


		sheets[1] = new SheetDTO();
		sheets[1].setSheetName("111111");
		sheets[1].setHtml(html.toString());

		TableListToXls.process(sheets,fout);
//		TableToXls.process(html, fout);
		fout.close();
	}
}
