package io.github.oliviercailloux.tutorial;

import java.util.HashMap;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.microsoft.graph.authentication.IAuthenticationProvider;
import com.microsoft.graph.requests.GraphServiceClient;

import okhttp3.Request;

public class App {
	@SuppressWarnings("unused")
	private static final Logger LOGGER = LoggerFactory.getLogger(Authenticator.class);

	public static void main(String[] args) {
		// Get an access token

		IAuthenticationProvider token = Authenticator.getAuthenticationProvider();

		GraphServiceClient<Request> graphClient = GraphServiceClient.builder().authenticationProvider(token)
				.buildClient();

		// Get the file Id of Book1.xlsx

		String fileId = WorksheetReader.getFileId("Book1.xlsx", graphClient);

		// In this section we will write in the sheet
		// we create a map that will contain the assignements. We will loop after on the
		// elements of the map to write in the sheet
		Map<Integer, Map<String, String>> assignementMapAll = new HashMap<>();
		Map<String, String> assignement1 = new HashMap<>();
		assignement1.put("level", "DE1");
		assignement1.put("Semestre", "1");
		assignement1.put("Course", "ALgébre");
		assignement1.put("Type", "CMTD");
		assignement1.put("NbH", "15");
		HashMap<String, String> assignement2 = new HashMap<>();
		assignement2.put("level", "DE1");
		assignement2.put("Semestre", "1");
		assignement2.put("Course", "Anglais");
		assignement2.put("Type", "TD");
		assignement2.put("NbH", "35");
		assignementMapAll.put(1, assignement1);
		assignementMapAll.put(2, assignement2);

		LOGGER.info("We are writing in the sheet ....");

		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 0, 0, "Assignement perTeacher", graphClient);

		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 2, 0, "FirstName", graphClient);
		WorksheetWriter.changeCellColor(fileId, "Sheet1", 2, 0, "#a2eaea", graphClient);
		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 2, 2, "Name", graphClient);
		WorksheetWriter.changeCellColor(fileId, "Sheet1", 2, 2, "#a2eaea", graphClient);
		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 3, 0, "Done", graphClient);

		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 3, 2, "Lopez", graphClient);

		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 5, 0, "Level", graphClient);
		WorksheetWriter.changeCellColor(fileId, "Sheet1", 5, 0, "#a2eaea", graphClient);
		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 5, 1, "Course", graphClient);
		WorksheetWriter.changeCellColor(fileId, "Sheet1", 5, 1, "#a2eaea", graphClient);
		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 5, 2, "Type", graphClient);
		WorksheetWriter.changeCellColor(fileId, "Sheet1", 5, 2, "#a2eaea", graphClient);
		WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", 5, 3, "NbH", graphClient);
		WorksheetWriter.changeCellColor(fileId, "Sheet1", 5, 3, "#a2eaea", graphClient);

		int row = 6;
		for (Map<String, String> assignement : assignementMapAll.values()) {

			WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", row, 0, assignement.get("level"), graphClient);
			WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", row, 1, assignement.get("Course"), graphClient);
			WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", row, 2, assignement.get("Type"), graphClient);
			WorksheetWriter.writeStringIntheSheet(fileId, "Sheet1", row, 3, assignement.get("NbH"), graphClient);
			row++;
		}

		LOGGER.info("done");

		// An example to read a cell value
		LOGGER.info("value of the cell A1 :" + WorksheetReader.getCellValue(fileId, "Sheet1", graphClient, 0, 0));

	}
}
