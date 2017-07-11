import java.util.Date;
import javax.activation.FileDataSource;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import java.text.SimpleDateFormat;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.BodyPart;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMultipart;
import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.Label;
import jxl.write.WriteException;







class mail
{
public static String s1,s2,nme;
//static WritableWorkbook wwbCopy;
	public mail(String s,String sc,String e,String a,String name)
	{
	s1=s;
	s2=sc;
	nme=name;
	String filename =a;
	final String SSL_FACTORY = "javax.net.ssl.SSLSocketFactory";
  	
     Properties props = System.getProperties();
     props.setProperty("mail.smtp.host", "smtp.gmail.com");
     
     props.setProperty("mail.smtp.socketFactory.fallback", "true");
     props.setProperty("mail.smtp.port", "587");
     props.setProperty("mail.smtp.socketFactory.port", "587");
     props.put("mail.smtp.auth", "true");
     props.put("mail.debug", "true");
     props.put("mail.store.protocol", "pop3");
     props.put("mail.transport.protocol", "smtp");
     props.put("mail.smtp.starttls.enable","true");
     props.put("mail.smtp.ssl.trust", "smtp.gmail.com");
     final String username = "chauhan.sourabh13@gmail.com";//
     final String password = "#your_password_here";
     try{
     Session session = Session.getDefaultInstance(props,  new Authenticator(){
                             protected PasswordAuthentication getPasswordAuthentication() {
                                return new PasswordAuthentication(username, password);
                             }});
   // -- Create a new message --
     Message msg = new MimeMessage(session);

          // -- Set the FROM and TO fields --
     msg.setFrom(new InternetAddress("chauhan.sourabh13@gmail.com"));
     msg.setRecipients(Message.RecipientType.TO, 
                      InternetAddress.parse(e,false));
	
	
	
	
	
	
	// code for sending the bcc
	
	
		/*	File inputWorkbook = new File("C:\\Users\\S\\Desktop\\db1.xls");       
            Workbook w;
           
            w = Workbook.getWorkbook(inputWorkbook);   
            wwbCopy = Workbook.createWorkbook(inputWorkbook, w);         
            // Get the first sheet            
            Sheet sheet = w.getSheet(0);
			
			Address ad[]=new Address[sheet.getRows()];
			
			for(int i=0;i<sheet.getRows();i++)
			{
				Cell cell = sheet.getCell(1, i); //               
				CellType type = cell.getType();
				String temp=cell.getContents();
				if(temp.equals(e)==false)
				{
				ad[i]=cell.getContents();
				}
			}
			
	
		
			msg.addRecipients(Message.RecipientType.BCC,ad);
	*/
	
	
	
/*	msg.addRecipient(Message.RecipientType.BCC, new InternetAddress(
            "ashok.cse14@nituk.ac.in"));
    msg.addRecipient(Message.RecipientType.BCC, new InternetAddress(
            "sgnsashok@gmail.com"));			
					  
	*/				  
	
	
	
	
	
	
	String bir="HAPPY BIRTHDAY TO:"+nme;
	String ani="HAPPY MARRAIGE ANNIVERSARY TO:"+nme;				  
	if(filename.equals("D:\\projects\\e-mail scheduler\\birth.jpg"))
	{		
		msg.setSubject(bir);
	}
	
	else
		msg.setSubject(ani);
	 
	 
	   // This mail has 2 part, the BODY and the embedded image
         MimeMultipart multipart = new MimeMultipart("related");

         // first part (the html)
         BodyPart messageBodyPart = new MimeBodyPart();
         String htmlText = s1+"<br/>"+s2+"<br/><br/></br>"+"<img src=\"cid:image\"></br></br></br><p><font color=\"green\">Thanks & Regards,<br/><br/>Sourabh Singh Chauhan<br/>B.tech Final Year(BT14CSE022),<br/>Computer Science & Engineering Department,<br/>National Institute of Technology, Uttarakhand,<br/>Uttarakhand - 246174, INDIA<br/>Mobile: +917895188090<br/>sourabh13.cse14@nituk.ac.in</font></p>";
         messageBodyPart.setContent(htmlText, "text/html");
         // add it
         multipart.addBodyPart(messageBodyPart);

         // second part (the image)
         messageBodyPart = new MimeBodyPart();
         DataSource fds = new FileDataSource(filename);

         messageBodyPart.setDataHandler(new DataHandler(fds));
         messageBodyPart.setHeader("Content-ID", "<image>");

         // add image to the multipart
         multipart.addBodyPart(messageBodyPart);

         // put everything together
         msg.setContent(multipart);
         // Send message
         Transport.send(msg);
	 
	 
	 
	 
    

/*	msg.setText(s1);
     BodyPart messageBodyPart = new MimeBodyPart();

         // Now set the actual message
         messageBodyPart.setText(s1);

         // Create a multipar message
         Multipart multipart = new MimeMultipart();

         // Set text message part
         multipart.addBodyPart(messageBodyPart);

         // Part two is attachment
         messageBodyPart = new MimeBodyPart();
         
         DataSource source = new FileDataSource(filename);
         messageBodyPart.setDataHandler(new DataHandler(source));
         messageBodyPart.setFileName(filename);
         multipart.addBodyPart(messageBodyPart);

         // Send the complete message parts
         msg.setContent(multipart);
     msg.setSentDate(new Date());
     Transport.send(msg);*/
	 
	 
	 
	 
	 
	 
     System.out.println("Message sent.");
  }catch (MessagingException ee){ System.out.println("Erreur d'envoi, cause: " + ee);}
  
	
}
}





class soft 
{
    static WritableWorkbook wwbCopy;
    public static void main(String[] args) 
	{
		
		
		try
		{
			
			int count=0,i;
            Calendar cal = Calendar.getInstance();
            SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
            System.out.println( sdf.format(cal.getTime()) );
            int day = cal.get(Calendar.DATE);
            int month = cal.get(Calendar.MONTH)+1;
            int year = cal.get(Calendar.YEAR);
            System.out.println(day);
            File inputWorkbook = new File("D:\\projects\\e-mail scheduler\\db1.xls");       
            Workbook w;
           
            w = Workbook.getWorkbook(inputWorkbook);   
            wwbCopy = Workbook.createWorkbook(inputWorkbook, w);         
            // Get the first sheet            
            Sheet sheet = w.getSheet(0);            
            // Loop over first 10 column and lines sheet.getRows()   
            System.out.println(sheet.getRows());        
			
			
			
			for (int tu = 0; tu<sheet.getRows() ; tu++) 
            {
				
				Cell cell = sheet.getCell(0, tu); //               
				CellType type = cell.getType();
				String ss=cell.getContents();
				System.out.print(""+ ss);            
				System.out.println(); 

				String arr[]= ss.split("/");
				int a=Integer.parseInt(arr[0]);
				int b=Integer.parseInt(arr[1]);
				System.out.println(year);
				System.out.println(a);
				System.out.println(b);
				System.out.println(month);
				System.out.println(day);
				cell = sheet.getCell(2, tu);//
				String rt=cell.getContents();
				int rs=Integer.parseInt(rt);
				System.out.println(tu);


				cell = sheet.getCell(3, tu);
                String last=cell.getContents();
				String bd="D:\\projects\\e-mail scheduler\\birth.jpg";
				String an="D:\\projects\\e-mail scheduler\\ann.PNG";
				
				cell = sheet.getCell(4, tu);
                String nm=cell.getContents();
				
				
				String t="Dear Sir/Ma"+"\'"+"am";
				
				 if((rs==year) && (a==month)&&(b==day)&& (last.equals("B"))==true )
				{
					cell = sheet.getCell(1, tu);//
					String e=cell.getContents();
					System.out.println("hi");
					
					mail m=new mail(t,"On behalf of Chauhan family, I wish you very very Happy Birthday",e,bd,nm);
					System.out.println("Message sent.");
					WritableSheet wshTemp = wwbCopy.getSheet(0);
					rs++;
					String rr=rs+"";
					Label label= new Label(2, tu, rr);
					try {
							wshTemp.addCell(label);
						} 
					catch (Exception ee) {}
				}
				
				
				
				
				if((rs==year) && (a==month)&&(b==day)&& (last.equals("A")==true ))
				{
					cell = sheet.getCell(1, tu);//
					String e=cell.getContents();
					System.out.println("hi");
					mail m=new mail(t,"On behalf of Chauhan family, I wish you very very Happy Marraige Anniversary.",e,an,nm);
					System.out.println("Message sent.");
					WritableSheet wshTemp = wwbCopy.getSheet(0);
					rs++;
					String rr=rs+"";
					Label label= new Label(2, tu, rr);
					try {
					wshTemp.addCell(label);
						} 
					catch (Exception ee) {}
				}
				
				
				
				
				if((rs==year) && (a==month)&&(b==day)&& (last.equals("AB")==true ))
				{
					cell = sheet.getCell(1, tu);//
					String e=cell.getContents();
					System.out.println("hi");
					mail m=new mail(t,"On behalf of Chauhan family, I wish you very very Happy Birthday and Marraige Anniversary.",e,bd,nm);
					System.out.println("Message sent.");
					WritableSheet wshTemp = wwbCopy.getSheet(0);
					rs++;
					String rr=rs+"";
					Label label= new Label(2, tu, rr);
					try {
					wshTemp.addCell(label);
						} 
					catch (Exception ee) {}
				}
				
				
				
				
				
				if((rs==year)&& (a==month)&&(b<day)&&(last.equals("A")==true))
				{
					cell = sheet.getCell(1, tu);//
					String e1=cell.getContents();
					System.out.println("hi");
					mail m=new mail(t,"On behalf of Chauhan family, I wish you belated Happy Marraige Anniversary.",e1,an,nm);
					System.out.println("Message sent.");
					WritableSheet wshTemp = wwbCopy.getSheet(0);
					rs++;
					String rr=rs+"";
					Label label= new Label(2, tu, rr);
					try {
					wshTemp.addCell(label);
					} 
					catch (Exception ee) {}  
				}
				
				
				
				
				
				if((rs==year)&& (a==month)&&(b<day)&&(last.equals("AB")==true))
				{
					cell = sheet.getCell(1, tu);//
					String e1=cell.getContents();
					System.out.println("hi");
					mail m=new mail(t,"On behalf of NIT Uttarakhand family, I wish you belated Happy Birthday & Marrauge Anniversary.",e1,bd,nm);
					System.out.println("Message sent.");
					WritableSheet wshTemp = wwbCopy.getSheet(0);
					rs++;
					String rr=rs+"";
					Label label= new Label(2, tu, rr);
					try {
					wshTemp.addCell(label);
						} 
					catch (Exception ee) {}  
				}
				
				
				
				
				
				if((rs==year)&& (a==month)&&(b<day)&&(last.equals("B")==true))
				{
					cell = sheet.getCell(1, tu);//
					String e1=cell.getContents();
					System.out.println("hi");
					mail m=new mail(t,"On behalf of Chauhan family, I wish you belated Happy Birthday.",e1,bd,nm);
					System.out.println("Message sent.");
					WritableSheet wshTemp = wwbCopy.getSheet(0);
					rs++;
					String rr=rs+"";
					Label label= new Label(2, tu, rr);
					try {
						wshTemp.addCell(label);
						} 
					catch (Exception ee) {}  
				}
				
				
				
			}
			
			
			
			 try {
					// Closing the writable work book
					wwbCopy.write();
					wwbCopy.close();
					w.close();
					// Closing the original work book
         
				} catch (Exception ee){} 
                
			
		}
		
		
		catch(Exception ee){}
		
		
	}
}
