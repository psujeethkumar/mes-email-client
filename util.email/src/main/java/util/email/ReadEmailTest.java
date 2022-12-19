package util.email;

import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.util.Map;

import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicHeader;

import com.fasterxml.jackson.databind.JavaType;
import com.fasterxml.jackson.databind.ObjectMapper;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ConnectingIdType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.misc.ImpersonatedUserId;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

public class ReadEmailTest {

	static String mailAddress = "fx.dev@luminus.be";
	static String tenentId = "c272261a-5184-4652-9972-a7bf2bf53f4d";
	static String clientId = "83afd42d-6d51-44ef-8dbf-4989a32af887";
	static String client_secret = "YHH8Q~OCdS_WQLvms0hW9UJ3AsMXIzDbn5ce7cHy";
	static final String uri = "https://outlook.office365.com/ews/exchange.asmx";
	static final String MAILBOX = "fx.dev@luminus.be";

	public static String getAuthToken(String tanantId, String clientId, String client_secret) throws ClientProtocolException, IOException {
		CloseableHttpClient client = HttpClients.createDefault();
		// HttpPost loginPost = new HttpPost("https://login.microsoftonline.com/" +
		// tanantId + "/oauth2/token");

		HttpPost loginPost = new HttpPost("https://login.microsoftonline.com/c272261a-5184-4652-9972-a7bf2bf53f4d/oauth2/v2.0/token");

		String scopes = "https://outlook.office365.com/.default";

		// String scopes = "";

		String encodedBody = "client_id=" + clientId + "&scope=" + scopes + "&client_secret=" + client_secret + "&grant_type=client_credentials";
		loginPost.setEntity(new StringEntity(encodedBody, ContentType.APPLICATION_FORM_URLENCODED));
		loginPost.addHeader(new BasicHeader("cache-control", "no-cache"));
		CloseableHttpResponse loginResponse = client.execute(loginPost);
		System.out.println("RESONSE CODE : " + loginResponse.getStatusLine().getStatusCode());
		InputStream inputStream = loginResponse.getEntity().getContent();
		byte[] response = inputStream.readAllBytes();
		ObjectMapper objectMapper = new ObjectMapper();
		JavaType type = objectMapper.constructType(objectMapper.getTypeFactory().constructParametricType(Map.class, String.class, String.class));
		Map<String, String> parsed = new ObjectMapper().readValue(response, type);
		return parsed.get("access_token");
	}

//	public static void main1(String[] args) {
//
//		Properties props = new Properties();
//
//		props.put("mail.store.protocol", "imap4");
//		props.put("mail.imap.host", "outlook.office365.com");
//		props.put("mail.imap.port", "993");
//		props.put("mail.imap.ssl.enable", "true");
//		props.put("mail.imap.starttls.enable", "true");
//		props.put("mail.imap.auth", "true");
//		props.put("mail.imap.auth.mechanisms", "XOAUTH2");
//		props.put("mail.imap.user", mailAddress);
//		props.put("mail.debug", "true");
//		props.put("mail.debug.auth", "true");
//		// props.put("mail.imap.auth.plain.disable", "true");
//		// props.setProperty("mail.imap.connectiontimeout", "5000");
//		// props.setProperty("mail.imap.timeout", "5000");
//
//		// open mailbox....
//		String token = "";
//		try {
//			token = getAuthToken(tenentId, clientId, client_secret);
//			System.out.println("TOKEN : " + token);
//		} catch (ClientProtocolException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		Session session = Session.getInstance(props);
//		session.setDebug(true);
//		Store store = null;
//		try {
//			store = session.getStore("imap");
//		} catch (NoSuchProviderException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		if (store != null) {
//			try {
//				store.connect("https://outlook.office365.com/.default", mailAddress, token);
//			} catch (MessagingException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//		}
//	}

	public static void main(String[] args) throws Exception {
		String token = "";
		ExchangeService exchangeService = null;
		try {
			token = getAuthToken(tenentId, clientId, client_secret);
			System.out.println("TOKEN : " + token);
			// String usersListFromGraph = getUsersListFromGraph(result.accessToken());
			// System.out.println("Users in the Tenant = " + usersListFromGraph);

			exchangeService = new ExchangeService();
			exchangeService.setUrl(new URI(uri));
			exchangeService.getHttpHeaders().put("Authorization", "Bearer " + token);

			exchangeService.setImpersonatedUserId(new ImpersonatedUserId(ConnectingIdType.SmtpAddress, MAILBOX));
			ItemView itemView = new ItemView(20);
			FindItemsResults<Item> findResults = exchangeService.findItems(WellKnownFolderName.Inbox, itemView);
			if (findResults != null) {
				exchangeService.loadPropertiesForItems(findResults, PropertySet.FirstClassProperties);
				System.out.println("found " + findResults.getTotalCount() + " mail items in mailbox " + MAILBOX);
				EmailMessage message = ((EmailMessage) findResults.getItems().get(0));
				System.out.println("BODY CONTENT : ");
				System.out.println(message.getBody().toString());
				message.delete(DeleteMode.SoftDelete);
			} else {
				System.out.println("found no mail items in mailbox " + MAILBOX);
			}

		} catch (Exception ex) {
			System.out.println("Oops! We have an exception of type - " + ex.getClass());
			System.out.println("Exception message - " + ex.getMessage());
			throw ex;
		} finally {
			exchangeService.close();
		}

	}
}
