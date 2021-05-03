package io.github.oliviercailloux;

import java.net.URL;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

import okhttp3.Request;

import com.azure.identity.DeviceCodeCredential;
import com.azure.identity.DeviceCodeCredentialBuilder;

import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.logger.DefaultLogger;
import com.microsoft.graph.logger.LoggerLevel;
import com.microsoft.graph.models.Attendee;
import com.microsoft.graph.models.DateTimeTimeZone;
import com.microsoft.graph.models.EmailAddress;
import com.microsoft.graph.models.Event;
import com.microsoft.graph.models.ItemBody;
import com.microsoft.graph.models.User;
import com.microsoft.graph.models.AttendeeType;
import com.microsoft.graph.models.BodyType;
import com.microsoft.graph.options.HeaderOption;
import com.microsoft.graph.options.Option;
import com.microsoft.graph.options.QueryOption;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.requests.EventCollectionPage;
import com.microsoft.graph.requests.EventCollectionRequestBuilder;

public class Graph {

	private static GraphServiceClient<Request> graphClient = null;
	private static TokenCredentialAuthProvider authProvider = null;

	public static void initializeGraphAuth(String applicationId, List<String> scopes) {
		// Create the auth provider
		final DeviceCodeCredential credential = new DeviceCodeCredentialBuilder().clientId(applicationId)
				.challengeConsumer(challenge -> System.out.println(challenge.getMessage())).build();

		authProvider = new TokenCredentialAuthProvider(scopes, credential);

		// Create default logger to only log errors
		DefaultLogger logger = new DefaultLogger();
		logger.setLoggingLevel(LoggerLevel.ERROR);

		// Build a Graph client
		graphClient = GraphServiceClient.builder().authenticationProvider(authProvider).logger(logger).buildClient();
	}

	public static String getUserAccessToken() {
		try {
			URL meUrl = new URL("https://graph.microsoft.com/v1.0/me");
			return authProvider.getAuthorizationTokenAsync(meUrl).get();
		} catch (Exception ex) {
			return null;
		}
	}
}