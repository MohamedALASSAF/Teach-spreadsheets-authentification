package io.github.oliviercailloux;
import com.azure.identity.InteractiveBrowserCredential;
import com.azure.identity.InteractiveBrowserCredentialBuilder;
import com.google.common.collect.ImmutableList;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.models.User;
import com.microsoft.graph.requests.GraphServiceClient;
import java.util.List;
import okhttp3.Request;

public class App 
{
	private final static String CLIENT_ID = "4b5d4470-571b-4602-816e-87cc364008e7";
	  private final static List<String> SCOPES = ImmutableList.of("User.Read");
	  

	  public static void main(String[] args) throws Exception {
	    final InteractiveBrowserCredential interactiveBrowserCredential =
	        new InteractiveBrowserCredentialBuilder().clientId(CLIENT_ID)
	            .redirectUrl("http://localhost/").build();
	    final TokenCredentialAuthProvider tokenCredentialAuthProvider =
	        new TokenCredentialAuthProvider(SCOPES, interactiveBrowserCredential);

	    final GraphServiceClient<Request> graphClient = GraphServiceClient.builder()
	        .authenticationProvider(tokenCredentialAuthProvider).buildClient();

	    final User me = graphClient.me().buildRequest().get();
	    System.out.println(me.mail);
	    
	  }
}
