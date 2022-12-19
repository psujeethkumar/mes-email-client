package util.email;

import java.net.MalformedURLException;
import java.net.URI;
import java.util.Calendar;
import java.util.Collections;
import java.util.concurrent.CompletableFuture;

import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ClientCredentialParameters;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import com.nimbusds.jose.JWSObject;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ConnectingIdType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.misc.ImpersonatedUserId;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

public class ReadEmail {

	// Get below values from the message flow properties.

	private static String TENENT_ID = "";
	private static String CLIENT_ID = "";
	private static String CLIENT_SECRET = "";
	private static String SCOPE = "";
	private static final String URI = "";
	private static final String AUTHORITY = "https://login.microsoftonline.com/" + TENENT_ID + "/authorize";
	private static final String MAILBOX = "";

	private static ConfidentialClientApplication app;

	private static IAuthenticationResult getAccessTokenByClientCredentialGrant() throws Exception {

		// With client credentials flows the scope is ALWAYS of the shape
		// "resource/.default", as the
		// application permissions need to be set statically (in the portal), and then
		// granted by a
		// tenant administrator
		ClientCredentialParameters clientCredentialParam = ClientCredentialParameters.builder(Collections.singleton(SCOPE)).build();

		CompletableFuture<IAuthenticationResult> future = app.acquireToken(clientCredentialParam);
		return future.get();
	}

	private static void BuildConfidentialClientObject() throws Exception {

		// Load properties file and set properties used throughout the sample
		try {
			app = ConfidentialClientApplication.builder(CLIENT_ID, ClientCredentialFactory.createFromSecret(CLIENT_SECRET)).authority(AUTHORITY).build();
		} catch (MalformedURLException e) {
			e.printStackTrace();
			throw new Exception("Invalid URl to connect to Miscrosoft Exchange : " + AUTHORITY + ". Check the tenant Id.");
		}
	}

	public static void main(String[] args) throws Exception {

		try {
			BuildConfidentialClientObject();
		} catch (Exception e) {
			e.printStackTrace();
			throw new Exception("Error while retrieving the OAuth token.");
		}
		IAuthenticationResult result = getAccessTokenByClientCredentialGrant();
		String token = result.accessToken();
		System.out.println(token);
		validateToken(token);

		
		ExchangeService exchangeService = new ExchangeService();
		exchangeService.setUrl(new URI(URI));
		// exchangeService.getHttpHeaders().put("Authorization", "Bearer " +
		// result.accessToken());
		exchangeService.getHttpHeaders().put("Authorization", token);
		exchangeService.setImpersonatedUserId(new ImpersonatedUserId(ConnectingIdType.SmtpAddress, MAILBOX));

		// Setting the item length to read 10 email at a time.

		ItemView itemView = new ItemView(1);
		FindItemsResults<Item> findResults = exchangeService.findItems(WellKnownFolderName.Inbox, itemView);

		// Check if there is any new email arrived. If yes, process one by one.
		if (findResults != null) {
			for (Item item : findResults) {
				exchangeService.loadPropertiesForItems(findResults, PropertySet.FirstClassProperties);
				System.out.println("found " + findResults.getTotalCount() + " mail items in mailbox " + MAILBOX);
				EmailMessage message = (EmailMessage) item;
				System.out.println("BODY CONTENT : ");
				System.out.println(message.getBody().toString());
				// IMPORTANT : Message must be soft deleted after it is stored in the PERS.
				// message.delete(DeleteMode.SoftDelete);
			}
		} else {
			// If there are no new email, stop the flow by propagating to appropriate
			// terminal.
			System.out.println("found no mail items in mailbox " + MAILBOX);
		}
		exchangeService.close();

	}

	private static boolean validateToken(String token) throws Exception {
		boolean expired = false;
		JWSObject jwsObject = JWSObject.parse(token);

		Long exp = (Long) jwsObject.getPayload().toJSONObject().get("exp") * 1000;

		Calendar currentTime = Calendar.getInstance();
		currentTime.add(Calendar.SECOND, -10);
		Long currentTimeInMilliSeconds = currentTime.getTimeInMillis();

		System.out.println("Expiring at : " + new java.util.Date(exp));

		System.out.println("Current time : " + new java.util.Date(currentTimeInMilliSeconds));

		if (exp <= currentTimeInMilliSeconds) {
			expired = true;
		}
		System.out.println(expired);
		return expired;
	}

}
