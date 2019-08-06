import java.util.Calendar;
import java.util.TimeZone;

import com.aspose.email.MailMessage;
import com.aspose.email.SaveOptions;

//w w  w  . j ava2 s  . c om
public class Main {

  public static void main(String args[]) {
   
	  
	  MailMessage msg=MailMessage.load("C:\\Users\\DELL\\Desktop\\sonika@sysinfotools.onmicrosoft.com_EMLFormat\\"
	  		+ "sonika@sysinfotools.onmicrosoft.com\\Inbox\\You've joined the DemoGroup group.eml");
	//get Calendar instance
	    Calendar now = Calendar.getInstance();
	    
	    //get current TimeZone using getTimeZone method of Calendar class
	    TimeZone timeZone = now.getTimeZone();
	    
	    //display current TimeZone using getDisplayName() method of TimeZone class
	    System.out.println("Current TimeZone is : " + timeZone.getRawOffset());
	    
	    msg.setTimeZoneOffset(timeZone.getRawOffset());
	    msg.save("C:\\Users\\DELL\\Desktop\\1.mht", SaveOptions.getDefaultMhtml());
  }
  
  
}