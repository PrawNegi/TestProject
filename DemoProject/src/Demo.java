import java.io.IOException;
import com.aspose.email.EWSClient;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.ExchangeFolderInfoCollection;
import com.aspose.email.ExchangeFolderType;
import com.aspose.email.ExchangeMessageInfo;
import com.aspose.email.ExchangeMessageInfoCollection;
import com.aspose.email.FolderInfo;
import com.aspose.email.IEWSClient;
import com.aspose.email.License;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiDistributionList;
import com.aspose.email.MapiMessage;
import com.aspose.email.MessageInfo;
import com.aspose.email.PersonalStorage;

public class Demo {
	
	static class D{
		static void listAllFolders() {
			
			License lic=new License();
			lic.setLicense("D:\\Aspose Workspace\\Aspose.Email (1).lic");
			
			IEWSClient client=null;
			try
			{
				String rootfolder="Root";
				client = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx", 
						"sonika@sysinfotools.onmicrosoft.com", "1qaz@WSX","");
				
				String inputPst = "E:\\Pst ,Mbox , Olm Samples\\Aspose Sample Files\\contact.pst";
				
				String path = inputPst;
					PersonalStorage personalStorage = PersonalStorage.fromFile(path);
					    
						
					        FolderInfo PSTFolder = personalStorage.getRootFolder().getSubFolder("Deleted Items");

					        int size=PSTFolder.getContentCount();
					        
					        for(MessageInfo messageInfo : PSTFolder.getContents())
					        {
					        	MapiMessage message= personalStorage.extractMessage(messageInfo);
					        	
					        	if(messageInfo.getMessageClass().equals("IPM.Note"))
					        	{
					        		
					        	}
					        	else if(messageInfo.getMessageClass().equals("IPM.Appointment"))
					        	{

					        	}
					        	else if(messageInfo.getMessageClass().equals("IPM.Contact"))
					        	{
					        		MapiContact mapiContact = (MapiContact)message.toMapiMessageItem();
					        	}
					        	else if(messageInfo.getMessageClass().equals("IPM.DistList"))
					        	{
					        		MapiDistributionList dlist=(MapiDistributionList) message.toMapiMessageItem();
					        	}


					        	// String mess=client.AppendMessage(client.MailboxInfo.InboxUri, message.ToMailMessage(new
					        	String mess = client.appendMessage(client.getMailboxInfo().getDeletedItemsUri(), message.toMailMessage(new MailConversionOptions()));
					        	
					        }
					        	 
//				String inbox = client.getMailboxInfo().getDeletedItemsUri();
//				String folderName1 = "EMAILNET-35054";
//				String subFolderName0 = "2015";
//				String folderName2 = folderName1 +  "/" +  subFolderName0;
//				ExchangeFolderInfo rootFolderInfo = null;
//				ExchangeFolderInfo folderInfo = null;
//			    try {
//			    	//MapiMessage mapim=MapiMessage.fromFile("C:\\Users\\DELL\\Desktop\\Eml2Pst_Pst2Msg_Format\\Top of Personal Folders\\Calendar\\Busy0.msg");
//			    	//MapiCalendar mapical=(MapiCalendar) mapim.toMapiMessageItem();
//			    	
//			    	//mapical.save("C:\\Users\\DELL\\Desktop\\Eml2Pst_Pst2Msg_Format\\Top of Personal Folders\\Calendar\\Busy0.ics", AppointmentSaveFormat.Ics);
//			    	//client.setUseSlashAsFolderSeparator(true);
//			    	//ExchangeFolderInfo exfolinfo=client.createFolder(client.getMailboxInfo().getRootUri(), folderName1);
//			    	Appointment app=Appointment.load("C:\\Users\\DELL\\Desktop\\Outlook_Pst2Eml_Format\\Top of Personal Folders\\Calendar\\Monthly Meeting - sent from win 7 0.ics");
//			    	if(app.getOrganizer()==null)
//			    	{
//			    		String str="Unknown";
//			    		MailAddress mailadd=MailAddress.to_MailAddress(str);
//			    		app.setOrganizer(mailadd);
//			    	}
//			    	System.out.println(app.getStartDate()+":"+app.getEndDate());
//			    	System.out.println(app.getOrganizer());
//			    	//ExchangeFolderInfo exfolinfo=client.createFolder(client.getMailboxInfo().getDeletedItemsUri(),"cal1", ExchangeFolderType.Appointment);
//			    	client.createAppointment(app,client.getMailboxInfo().getDeletedItemsUri());
//			    	//client.createFolder(client.getMailboxInfo().getCalendarUri(), folderName2);
//			    } finally {
////				    ExchangeFolderInfo subfolderInfo[] = new ExchangeFolderInfo[] { null };
////				    boolean outRefCondition2 = client.folderExists(inbox, folderName1, subfolderInfo);
////			        rootFolderInfo = subfolderInfo[0];
////			        Appointment app=Appointment.load("C:\\Users\\DELL\\AppData\\Local\\Temp\\tempcal.ics");
////			        client.createAppointment(app);
////			        if (outRefCondition2) {
////			        	ExchangeFolderInfo[] referenceToFolderInfo = { folderInfo };
////			        	boolean outRefCondition3 = client.folderExists(inbox, folderName2, /*out*/ referenceToFolderInfo);
////			        	folderInfo = referenceToFolderInfo[0];
////			        	if (outRefCondition3) {
////			        		//client.deleteFolder(folderInfo.getUri(), true);
////			    			//client.deleteFolder(rootFolderInfo.getUri(), true);
////			    		}
////			  		}
//			    }
			} finally {
				if (client != null)
			          client.dispose();
			}
			
			System.out.println(ExchangeFolderType.Undefined);
			System.out.println(ExchangeFolderType.Note);
			//System.out.println(ExchangeFolderType.RSSFeeds);
			System.out.println(ExchangeFolderType.Appointment);
			System.out.println(ExchangeFolderType.Contact);
			//System.out.println(ExchangeFolderType.QuickContacts);
			//System.out.println(ExchangeFolderType.ImContactsList);
			//System.out.println(ExchangeFolderType.DocumentLibraries);
			System.out.println(ExchangeFolderType.Journal);
			System.out.println(ExchangeFolderType.StickyNote);
			System.out.println(ExchangeFolderType.Task);
			System.out.println(ExchangeFolderType.Imap);
			//System.out.println(ExchangeFolderType.Unknown);
		}
	}
	
	private final static void listSubFolders(IEWSClient client, ExchangeFolderInfo folderInfo,String rootFolder) {
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
			System.out.println(ex.getMessage());
		}
	}
	
	public static void main(String[] args) throws IOException {
		Demo dm=new Demo();
		Demo.D.listAllFolders();
	}
}
