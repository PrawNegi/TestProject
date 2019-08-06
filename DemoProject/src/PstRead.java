import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import com.aspose.email.MapiMessage;
import com.aspose.email.OlmFolder;
import com.aspose.email.OlmStorage;

public class PstRead {
	
	public static void main(String[] args) throws IOException {
		OlmStorage storage = new OlmStorage("D:\\Olm samples\\Liveintent files.olm");
		PrintPath(storage, storage.getFolderHierarchy());
		
//		JFileChooser chooser = new JFileChooser();
//	    chooser.setCurrentDirectory(new File("."));
//	    chooser.setApproveButtonText("Save");
//	    chooser.setDialogTitle("Save");
//	    FileNameExtensionFilter filter = new FileNameExtensionFilter("csv File",
//	            new String[] { "csv" });
//	    chooser.setFileFilter(filter);
//	    chooser.addChoosableFileFilter(filter);
//	    //int r = chooser.showSaveDialog(null);
//	    File xmlFile=null;
//	    while (chooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION)
//	    {
//	    	xmlFile = new File(chooser.getSelectedFile().toString());
//			if (xmlFile.exists()) {
//			    int response = JOptionPane.showConfirmDialog(chooser, //
//			            "Do you want to replace the existing file?", //
//			            "Confirm", JOptionPane.YES_NO_OPTION, //
//			            JOptionPane.QUESTION_MESSAGE);
//			    if (response != JOptionPane.YES_OPTION) {
//			       continue;
//			    }
//			    else
//			    {
//			    	String xmlFilestr=xmlFile.getAbsolutePath();
//					xmlFile=new File(xmlFilestr);
//					xmlFile.delete();
//					System.out.println(xmlFile.createNewFile());
//					System.out.println(xmlFile.getAbsolutePath());
//					break;
//			    }
//			}
//			else
//			{
//				String xmlFilestr=xmlFile.getAbsolutePath()+".csv";
//				xmlFile=new File(xmlFilestr);
//				xmlFile.createNewFile();
//				System.out.println(xmlFile.getAbsolutePath());
//				break;
//			}
//	        // do the rest of your code
//	    }
	    
	}
	
	public static void PrintPath(OlmStorage storage, List<OlmFolder> folders) {
		for (OlmFolder folder : folders) {
			
			try
			{
				// print the current folder path
				System.out.println(folder.getPath());
				
				Iterator<MapiMessage> Mapimsg=storage.enumerateMessages(folder).iterator();
				int i=0;
				while(Mapimsg.hasNext())
				{
					System.out.println(Mapimsg.next().getMessageClass());
					i++;
				}
				
				System.out.println("Total Count : "+i);
				
				if (folder.getSubFolders().size() > 0) {
					PrintPath(storage, folder.getSubFolders());
				}
			}
			catch(Exception ep)
			{
				System.out.println(ep.getMessage());
			}
		
		}
	}
}
