import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.util.Properties;

public class sendMail {
    public static String myEmailAccount = "wyzh2002@163.com";
    public static String myEmailPassword = "wyzh2002";
    public static String myEmailSMTPHost = "smtp.163.com";

    public static void send(String imageDest, String reciever, String receiverName) throws Exception {
        Properties props = new Properties();
        props.setProperty("mail.transport.protocol", "smtp");
        props.setProperty("mail.smtp.host", myEmailSMTPHost);
        props.setProperty("mail.smtp.auth", "true");
        Session session = Session.getInstance(props);
        MimeMessage message = new MimeMessage(session);
        Transport tp = session.getTransport();
        tp.connect(myEmailAccount, myEmailPassword);
        message.setFrom(new InternetAddress(myEmailAccount,"Jacob","GBK"));
        message.setRecipient(MimeMessage.RecipientType.TO, new InternetAddress(reciever, receiverName, "GBK"));
        message.setSubject("促销明细", "GBK");

        MimeBodyPart image = new MimeBodyPart();
        DataHandler dh = new DataHandler(new FileDataSource(imageDest));
        image.setDataHandler(dh);
        image.setContentID("sheetImage");
        MimeBodyPart text = new MimeBodyPart();
        text.setContent("如下为本月待贵司确认的促销明细，如核对无误，请回复邮件确认，财务会按此结算，谢谢" +
                "<br/><img src='cid:sheetImage'/>", "text/html;charset=UTF-8");

        MimeMultipart text_image = new MimeMultipart();
        text_image.addBodyPart(text);
        text_image.addBodyPart(image);
        text_image.setSubType("related");
        message.setContent(text_image);
        message.saveChanges();
        tp.sendMessage(message, message.getAllRecipients());
        tp.close();
    }
}
