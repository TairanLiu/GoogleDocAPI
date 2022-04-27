import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.docs.v1.Docs;
import com.google.api.services.docs.v1.DocsScopes;
import com.google.api.services.docs.v1.model.*;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

    /* class to demonstarte use of Docs get documents API */
    public class DocsQuickstart {
        /** Application name. */
        private static final String APPLICATION_NAME = "Google Docs API Java Quickstart";
        /** Global instance of the JSON factory. */
        private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
        /** Directory to store authorization tokens for this application. */
        private static final String TOKENS_DIRECTORY_PATH = "tokens";
        private static String DOCUMENT_ID = "195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE";

        /**
         * Global instance of the scopes required by this quickstart.
         * If modifying these scopes, delete your previously saved tokens/ folder.
         */
        private static final List<String> SCOPES = Collections.singletonList(DocsScopes.DOCUMENTS);
        private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

        /**
         * Creates an authorized Credential object.
         * @param HTTP_TRANSPORT The network HTTP Transport.
         * @return An authorized Credential object.
         * @throws IOException If the credentials.json file cannot be found.
         */
        private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
            // Load client secrets.
            InputStream in = DocsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
            if (in == null) {
                throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
            }
            GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

            // Build flow and trigger user authorization request.
            GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                    HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                    .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                    .setAccessType("offline")
                    .build();
            LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
            Credential credential = new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
            //returns an authorized Credential object.
            return credential;
        }

        public static void generateDoc(String titleOfDoc) throws GeneralSecurityException, IOException {
            final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            Docs service = new Docs.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                    .setApplicationName(APPLICATION_NAME)
                    .build();

            // Prints the title of the requested doc:
            // https://docs.google.com/document/d/195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE/edit
            Document response = service.documents().get(DOCUMENT_ID).execute();
            String title = response.getTitle();

            //System.out.printf("The title of the doc is: %s\n", title);
            Document doc = new Document()
                    .setTitle(titleOfDoc);
            doc = service.documents().create(doc)
                    .execute();
            DOCUMENT_ID = doc.getDocumentId();
            System.out.println("create document with title: " + title);
        }
        public static void insertTitle(String docID) throws GeneralSecurityException, IOException {
            final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            Docs service = new Docs.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                    .setApplicationName(APPLICATION_NAME)
                    .build();
            List<Request> requests = new ArrayList<>();
            requests.add(new Request().setInsertText(new InsertTextRequest()
                    .setText("                                                                                         ")
                    .setLocation(new Location().setIndex(1))));
            requests.add(new Request().setInsertText(new InsertTextRequest()
                    .setText(docID+"\n")
                    .setLocation(new Location().setIndex(1))));
            requests.add(new Request()
                    .setUpdateTextStyle(new UpdateTextStyleRequest()
                            .setRange(new Range()
                                    .setStartIndex(1)
                                    .setEndIndex(docID.length()+10))
                            .setTextStyle(new TextStyle()
                                    .setFontSize(new Dimension()
                                            .setMagnitude(104.0)
                                            .setUnit("PT")))
                            .setFields("fontSize")));
            /*
            requests.add(new Request().setInsertText(new InsertTextRequest()
                    .setText("jiejie\n")
                    .setLocation(new Location().setIndex(7))));
            requests.add(new Request().setInsertText(new InsertTextRequest()
                    .setText("jiejie")
                    .setLocation(new Location().setIndex(15))));*/


            BatchUpdateDocumentRequest body = new BatchUpdateDocumentRequest().setRequests(requests);
            BatchUpdateDocumentResponse responses = service.documents()
                    .batchUpdate(DOCUMENT_ID, body).execute();
        }
        public static void insertText(String boldPart, String bodyPart) throws GeneralSecurityException, IOException {
            //System.out.println("reached");
            final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            Docs service = new Docs.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                    .setApplicationName(APPLICATION_NAME)
                    .build();
            List<Request> requests = new ArrayList<>();
            requests.add(new Request().setInsertText(new InsertTextRequest()
                    .setText("                                                                                         ")
                    .setLocation(new Location().setIndex(21))));
            requests.add(new Request().setInsertText(new InsertTextRequest()
                    .setText(boldPart+"\n")
                    .setLocation(new Location().setIndex(21))));
            requests.add(new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                    .setTextStyle(new TextStyle()
                            .setBold(true)
                            .setItalic(true))
                    .setRange(new Range()
                            .setStartIndex(21)
                            .setEndIndex(boldPart.length()+21))
                    .setFields("bold")));
            requests.add(new Request().setInsertText(new InsertTextRequest()
                    .setText(bodyPart)
                    .setLocation(new Location().setIndex(boldPart.length()+21+1))));
            BatchUpdateDocumentRequest body = new BatchUpdateDocumentRequest().setRequests(requests);
            BatchUpdateDocumentResponse response = service.documents()
                    .batchUpdate(DOCUMENT_ID, body).execute();
            //System.out.println("complete");
        }
        public static void insertImage(String docID) throws GeneralSecurityException, IOException {
            final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            Docs service = new Docs.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                    .setApplicationName(APPLICATION_NAME)
                    .build();
            List<Request> requests = new ArrayList<>();
            requests.add(new Request().setInsertInlineImage(new InsertInlineImageRequest()
                    .setUri(docID)
                    .setLocation(new Location().setIndex(60))
                    .setObjectSize(new Size()
                            .setHeight(new Dimension()
                                    .setMagnitude(500.0)
                                    .setUnit("PT"))
                            .setWidth(new Dimension()
                                    .setMagnitude(500.0)
                                    .setUnit("PT")))));

            BatchUpdateDocumentRequest body = new BatchUpdateDocumentRequest().setRequests(requests);
            BatchUpdateDocumentResponse response = service.documents()
                    .batchUpdate(DOCUMENT_ID, body).execute();
        }


        public static void main(String... args) throws IOException, GeneralSecurityException {
            generateDoc("Puppie");//ok
            insertTitle("Puppie");//ok
            insertText("puppie:","yaowen");
            insertImage("https://c-ssl.duitang.com/uploads/blog/202102/26/20210226130626_66a88.thumb.700_0.jpg");//ok

            ///
            /**/

        }
    }