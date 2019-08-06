import java.io.IOException;
import com.aspose.email.EWSClient;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.ExchangeFolderInfoCollection;
import com.aspose.email.ExchangeFolderType;
import com.aspose.email.ExchangeMessageInfo;
import com.aspose.email.ExchangeMessageInfoCollection;
import com.aspose.email.IEWSClient;
import com.aspose.email.MailMessage;

public class cs {
	
	private static void listAllFolders() {
		try
		{
			String rootfolder="Root";
			IEWSClient client = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx", "Sysdemo12@hotmail.com", "Sys@12345","");
			String rootUri = client.getMailboxInfo().getRootUri();
			client.setTimeout(Integer.MAX_VALUE);
			System.out.println("Downloading all messages from Inbox....");
			//ExchangeMailboxInfo mailboxInfo = client.getMailboxInfo();
			//System.out.println("Mailbox URI: " + mailboxInfo.getMailboxUri());
			
			// List all the folders from Exchange server
			ExchangeFolderInfoCollection folderInfoCollection = client.listSubFolders(rootUri);
			int size=folderInfoCollection.size();
			
			for (int i=0;i<size;i++) 
			{
				ExchangeFolderInfo folderInfo=folderInfoCollection.get_Item(i);
				// Call the recursive method to read messages and get sub-folders
				//System.out.println(folderInfo.getDisplayName());
				//listSubFolders(client, folderInfo,rootfolder);
			}
			
			System.out.println("All messages downloaded.");
		}
		catch (Exception ex) {
			System.out.println(ex.fillInStackTrace());
		}
		
		System.out.println(ExchangeFolderType.Appointment);
		System.out.println(ExchangeFolderType.Contact);
		System.out.println(ExchangeFolderType.Imap);
		System.out.println(ExchangeFolderType.Journal);
		System.out.println(ExchangeFolderType.Note);
		System.out.println(ExchangeFolderType.StickyNote);
		System.out.println(ExchangeFolderType.Task);
		System.out.println(ExchangeFolderType.Undefined);
	}

	private static void listSubFolders(IEWSClient client, ExchangeFolderInfo folderInfo,String rootFolder) {
		try
		{
			// Create the folder in disk (same name as on IMAP server)
			String path=rootFolder+"\\"+folderInfo.getDisplayName();
			System.out.println(path);
			// If this folder has sub-folders, call this method recursively to get messages
			ExchangeFolderInfoCollection folderInfoCollection = client.listSubFolders(folderInfo.getUri());
			
			ExchangeMessageInfoCollection msgInfoColl = client.listMessages(folderInfo.getUri());
			int msgsize=msgInfoColl.size();
			System.out.println(msgsize);
			
			for(int i=0;i<msgsize;i++)
			{
				ExchangeMessageInfo exmsg=msgInfoColl.get_Item(i);
				
				MailMessage msg=client.fetchMessage(exmsg.getUniqueUri());
				System.out.println(msg.getSubject());
				System.out.println(msg.getFrom());
				System.out.println("----------------------------------");
			}
			
			int size=folderInfoCollection.size();
			
			if(size>0)
			{
				//System.out.println("sub folders-");
				for (int i=0;i<size;i++)
				{
					ExchangeFolderInfo subfolderInfo=folderInfoCollection.get_Item(i);
					listSubFolders(client,subfolderInfo,path);
				}
			}
		}
		catch (Exception ex) 
		{
			System.out.println(ex.fillInStackTrace());
		}
	}
	
	public static void main(String[] args) throws IOException {
		listAllFolders();
	}
}
