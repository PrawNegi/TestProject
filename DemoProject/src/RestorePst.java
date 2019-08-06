import java.util.concurrent.atomic.AtomicInteger;

import com.aspose.email.AsposeException;
import com.aspose.email.BeforeItemCallback;
import com.aspose.email.EWSClient;
import com.aspose.email.ExchangeDistributionList;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.ExchangeFolderType;
import com.aspose.email.IEWSClient;
import com.aspose.email.ItemCallbackArgs;
import com.aspose.email.MailAddress;
import com.aspose.email.MailAddressCollection;
import com.aspose.email.MailMessage;
import com.aspose.email.MailMessageSaveType;
import com.aspose.email.MapiDistributionListMember;
import com.aspose.email.MsgSaveOptions;
import com.aspose.email.PersonalStorage;
import com.aspose.email.RestoreSettings;
import com.aspose.email.SaveOptions;

public class RestorePst {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
//		PersonalStorage pst = PersonalStorage.fromFile("E:\\Pst ,Mbox , Olm Samples\\Aspose Sample Files\\Eml2Pst.pst");
//		
//		IEWSClient client = EWSClient.getEWSClient("https://outlook.office365.com/ews/exchange.asmx", "sonika@sysinfotools.onmicrosoft.com", "1qaz@WSX");
//				System.out.println(client.getMailboxInfo().getDeletedItemsUri());
//				
//			System.out.println(client.getFolderInfo(client.getMailboxInfo().getDeletedItemsUri()).getFolderType());
//		ExchangeDistributionList exdistributionlist=new ExchangeDistributionList();
//		
//		exdistributionlist.setDisplayName("Demo name");
//		MailAddressCollection mailaddresscoll=new MailAddressCollection();
//		
//		MailAddress mailadd=MailAddress.to_MailAddress("sas@gmail.com");
//		mailadd.setAddress("bdc@gmail.com");
//		mailadd.setDisplayName("bdc");
//		mailaddresscoll.addItem(mailadd);
//		
//		client.createDistributionList(exdistributionlist, mailaddresscoll);
//		
//		ExchangeDistributionList[] distributionLists = client.listDistributionLists();
//		for(ExchangeDistributionList distributionList : distributionLists)
//		{
//		    MailAddressCollection members = client.fetchDistributionList(distributionList);
//		    for(MailAddress member : members)
//		    {
//		        System.out.println(member);
//		    }
//		}
		
		MailMessage eml = MailMessage.load("C:\\Users\\DELL\\Desktop\\sonika@sysinfotools.onmicrosoft.com_EMLFormat\\"
				+ "sonika@sysinfotools.onmicrosoft.com\\Inbox\\A change has been made to your subscript(1).eml");
		//Save the Email message to disk in Unicode format
		MsgSaveOptions msgSaveOptions = new MsgSaveOptions(MailMessageSaveType.getOutlookMessageFormatUnicode());
		msgSaveOptions.setPreserveOriginalDates(true);
		eml.save("C:\\Users\\DELL\\Desktop\\"+"LoadingEMLSavingToMSG_out.msg", msgSaveOptions);
	}
}
