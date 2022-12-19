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

		String tempToken = "Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IklkdUZKbVdfY09lS0xCRkxoLVBkS1JlRUpIUjREaXotYlNuYjluZGRhREkiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiJodHRwczovL291dGxvb2sub2ZmaWNlMzY1LmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2MyNzIyNjFhLTUxODQtNDY1Mi05OTcyLWE3YmYyYmY1M2Y0ZC8iLCJpYXQiOjE2NzE0NDA5NDUsIm5iZiI6MTY3MTQ0MDk0NSwiZXhwIjoxNjcxNDQ0ODQ1LCJhaW8iOiJFMlpnWU5qb3JQSHdkdStwRzNzQ2YvNXJYaWc1RHdBPSIsImFwcF9kaXNwbGF5bmFtZSI6Ik9BdXRoIE1haWwgYXBwIiwiYXBwaWQiOiI4M2FmZDQyZC02ZDUxLTQ0ZWYtOGRiZi00OTg5YTMyYWY4ODciLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9jMjcyMjYxYS01MTg0LTQ2NTItOTk3Mi1hN2JmMmJmNTNmNGQvIiwib2lkIjoiZmJjMDAwZmQtYTcxMC00MGM1LTk4YTYtOGU5MzU2Y2ZkNDAxIiwicmgiOiIwLkFVZ0FHaVp5d29SUlVrYVpjcWVfS19VX1RRSUFBQUFBQVBFUHpnQUFBQUFBQUFCSUFBQS4iLCJyb2xlcyI6WyJmdWxsX2FjY2Vzc19hc19hcHAiLCJNYWlsLlJlYWRXcml0ZSIsIk1haWwuUmVhZCIsIklNQVAuQWNjZXNzQXNBcHAiXSwic2lkIjoiODIxZjk3NmQtOWY3Yi00ZTQ1LWI0NzQtNDYyYjE0MmYzYjU3Iiwic3ViIjoiZmJjMDAwZmQtYTcxMC00MGM1LTk4YTYtOGU5MzU2Y2ZkNDAxIiwidGlkIjoiYzI3MjI2MWEtNTE4NC00NjUyLTk5NzItYTdiZjJiZjUzZjRkIiwidXRpIjoiNXN2S2dqTlRia3FZOWwtZnZ6cnNBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiMDk5N2ExZDAtMGQxZC00YWNiLWI0MDgtZDVjYTczMTIxZTkwIl19.ll8FuHN2Iw7qcYBACAUdE62ZD3gap2ViNj5PFLCityAz2av46LLxEszc2LhrF4bDJNKgrPSXujPNPjGmJbvWbjo06menBja6svq2Qe_8HK_v8-w30uAJhUcrkgI14YJlv4SvZUO85IsFjezeY4j_3P_6kK6E0LTshSJz9JsNy-klxrL10N8QA37U2oAxlbuCeYoS6IR12RgkyTJQZ04c3LCIJUb3aC7Nr2iqjMjgyW5KySoZyULjnIeX19u5JJMgZIA3uAk_mar10Ag5PLAoPptFx0A9-Q33Y4evFYhR8qEwmAuTexvGsEJ_xhuSlJVsoHsxp5zYje1_ia57sA-3uw";

		ExchangeService exchangeService = new ExchangeService();
		exchangeService.setUrl(new URI(URI));
		// exchangeService.getHttpHeaders().put("Authorization", "Bearer " +
		// result.accessToken());
		exchangeService.getHttpHeaders().put("Authorization", tempToken);
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
		token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6InJLSmYta2FpVi1iVmpNb1l0Q2E3TkRaQ1hpSjVZLVhfNkRrV1hTaE9wTkUiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiJodHRwczovL291dGxvb2sub2ZmaWNlMzY1LmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2MyNzIyNjFhLTUxODQtNDY1Mi05OTcyLWE3YmYyYmY1M2Y0ZC8iLCJpYXQiOjE2NzEzNTMzMTMsIm5iZiI6MTY3MTM1MzMxMywiZXhwIjoxNjcxMzU3MjEzLCJhaW8iOiJFMlpnWUtoOStWN2hTa1p5aE9iV3JKVDVVN0wyQXdBPSIsImFwcF9kaXNwbGF5bmFtZSI6Ik9BdXRoIE1haWwgYXBwIiwiYXBwaWQiOiI4M2FmZDQyZC02ZDUxLTQ0ZWYtOGRiZi00OTg5YTMyYWY4ODciLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9jMjcyMjYxYS01MTg0LTQ2NTItOTk3Mi1hN2JmMmJmNTNmNGQvIiwib2lkIjoiZmJjMDAwZmQtYTcxMC00MGM1LTk4YTYtOGU5MzU2Y2ZkNDAxIiwicmgiOiIwLkFVZ0FHaVp5d29SUlVrYVpjcWVfS19VX1RRSUFBQUFBQVBFUHpnQUFBQUFBQUFCSUFBQS4iLCJyb2xlcyI6WyJmdWxsX2FjY2Vzc19hc19hcHAiLCJNYWlsLlJlYWRXcml0ZSIsIk1haWwuUmVhZCIsIklNQVAuQWNjZXNzQXNBcHAiXSwic2lkIjoiNTYxZTllMTEtOWYwZS00NGQxLWFjNjMtODViYTNlNmU1NjZhIiwic3ViIjoiZmJjMDAwZmQtYTcxMC00MGM1LTk4YTYtOGU5MzU2Y2ZkNDAxIiwidGlkIjoiYzI3MjI2MWEtNTE4NC00NjUyLTk5NzItYTdiZjJiZjUzZjRkIiwidXRpIjoiLUJnZVZjaU8xMFduRHB3UHNOYmJBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiMDk5N2ExZDAtMGQxZC00YWNiLWI0MDgtZDVjYTczMTIxZTkwIl19.KWXkdH7IiT-sdwADOu9l1bi2IEIXfTQccg00CzZ1Ofjb_H4Elk6oh8kxfKOqGqpUd-rJw2lz3SwgE5OVjE582kolVn4muO4sI6wsUegAcZwEITIQ4fdeAVjZDPTaqXT1wHLgoGXPr5Xj6vQA0B35XRAR2jiNkmQ72soepgZelcuWRGBc2qSgXK7eG4CWi7D3Gfi9TgkDOIy2VqZJcQdJJyWSYTF3QGNr8VrxxlWx4wXzlBfeWdfVqT2exk36KXLJzBfiUMAT0QeT7bmBbQuK2bvQ_ROkJJj_HeVaNAO0JldFldiD9Yew-wpxItCun4_hrlwUo8KY4UDXETFUrNngtg";
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
