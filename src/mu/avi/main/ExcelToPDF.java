package mu.avi.main;

import java.util.Objects;

import mu.avi.controller.Controller;

public class ExcelToPDF {

	public static final String XLS_EXTENSION = "xls";
	public static final String XLSX_EXTENSION = "xlsx";

	public static void main(String[] args) {
		Controller controller = new Controller();
		try {
			Objects.requireNonNull(args);
			int count = 0;
			if (args.length == 0) {
				System.out.println("Wrong argument passed");
			} else {
				for (int i = 0; i < args.length; i++) {
					switch (args[i]) {
					case "-p": // set password
						i++;
						System.out.println("Password is: " + args[i]);
						break;
					default:
						boolean isXLS = controller.checkExcelFormat(args[i]);

						if (isXLS) {
							System.out.println("File extension is: " + XLS_EXTENSION);
							controller.readFromXLSExcel(args[i]);
						} else {
							System.out.println("File extension is: " + XLSX_EXTENSION);
							controller.readFromXLSXExcel(args[i]);
						}
						count++;

						System.out.println("file number: " + count + "finished");
					}
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("Wrong argument passed");
		}
	}

}
