package io.github.oliviercailloux;

import java.util.List;

import com.azure.identity.InteractiveBrowserCredential;
import com.azure.identity.InteractiveBrowserCredentialBuilder;
import com.google.common.collect.ImmutableList;
import com.google.gson.Gson;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.models.DriveSearchParameterSet;
import com.microsoft.graph.requests.DriveSearchCollectionPage;
import com.microsoft.graph.requests.GraphServiceClient;

import okhttp3.Request;

public class InteractiveFlow {
	private final static String CLIENT_ID = "afd352ef-e48a-4244-9340-95bfc83ef33c";
	private final static List<String> SCOPES = ImmutableList.of("User.Read", "Files.ReadWrite");

	public static void main(String[] args) throws Exception {

		final InteractiveBrowserCredential interactiveBrowserCredential = new InteractiveBrowserCredentialBuilder()
				.tenantId("common").clientId(CLIENT_ID).redirectUrl("http://localhost/").build();
		final TokenCredentialAuthProvider tokenCredentialAuthProvider = new TokenCredentialAuthProvider(SCOPES,
				interactiveBrowserCredential);

		final GraphServiceClient<Request> graphClient = GraphServiceClient.builder()
				.authenticationProvider(tokenCredentialAuthProvider).buildClient();

		/*
		 * DriveItem driveItem = new DriveItem(); driveItem.name = "New Folder"; Folder
		 * folder = new Folder(); driveItem.folder = folder;
		 * driveItem.additionalDataManager().put("@microsoft.graph.conflictBehavior",
		 * new JsonPrimitive("rename"));
		 */

		DriveSearchCollectionPage search = graphClient.me().drive()
				.search(DriveSearchParameterSet.newBuilder().withQ("Book1.xlsx").build()).buildRequest().get();
		Gson gson = new Gson();
		String json = gson.toJson(search);
		System.out.println(json);

		/*
		 * final User me = graphClient.me().buildRequest().get();
		 * System.out.println(me.displayName); System.out.println("done");
		 */
	}
}
