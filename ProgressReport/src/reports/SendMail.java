package reports;

import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

/*
 * *@author Vinoth N
 * Program to send to bulk emails
 */

public class SendMail {
	static String name, to;
	static Multipart multipart;
	public SendMail() {
		// TODO Auto-generated constructor stub
	}
	public SendMail(String name, String to,Multipart multipart){
		this.name=name;
		this.to=to;
		this.multipart=multipart;
	}
	/*
	 * In this class we have two methods:
	 * i. main method - main task will be executed
	 * ii. mailHost - returns the host object of a Mail class, helps to choose the mail host
	 */
	
	public void mail(){
		Mail mail=new Gmail();
		mail.send(name, to, multipart);
	}
}


abstract class Mail{
	/*
	 * It is an abstract contains 
	 * i. abstract method session() - to get the session of the user's chosen host
	 * ii. message() method - to insert subject, from address, and main content of the message and send it.
	 * iii. to() method - extracts to addresses from a text file and returns it as a property
	 * iv. send() method - used to manage the email. It will send organize to send email to each of the to addresses
	 */
	String from,password;
	abstract protected Session authenticate();//abstract method to choose the session for the host as per user wish
	
	private void message(String toperson, String toaddress,Multipart multipart) {
		/*
		 * Composition of message and regarding task is done in this method
		 */
		try {
			//MimeMessage object created with the help of host session
			MimeMessage message=new MimeMessage(authenticate());
			
			//Sender address is initialized using setFrom method
			message.setFrom(new InternetAddress(from));
			
			//Recipient address is set by setRecipient method
			message.setRecipient(Message.RecipientType.TO, new InternetAddress(toaddress));
			
			//Sets subject of the mail using setSubject method
			message.setSubject("Report Card");
			
			message.setContent(multipart);
			//Mail will be send
			Transport.send(message);
			System.out.println("Message successfully sent to "+toaddress);
			
		}catch(MessagingException ex) {//When transport object fails to send mail catch block will be executed
//			ex.printStackTrace();//Helps to trace the error
			System.out.println("Could not send the mail");
		}
	}
	public void send(String name, String to, Multipart multipart) {
		message(name, to, multipart);
	}
}


class Gmail extends Mail{
	/*
	 * Subclass of Mail class
	 * Implements authenticate() method - which returns the session for gmail
	 * Constructor assigns the from address and password for the session to get authenticate
	 */
	public Gmail() {//Gmail constructor to initialize the from and password
		from=System.getenv("USER_NAME");
		password=System.getenv("PASSWORD");
	}
	@Override
	protected Session authenticate() {//authenticate method to return session for gmail host
		String host="smtp.gmail.com";
		Properties prop=System.getProperties();
		prop.put("mail.smtp.host",host);
		prop.put("mail.smtp.port", "465");
		prop.put("mail.smtp.ssl.enable", "true");
		prop.put("mail.smtp.auth", "true");
		
		Session session =Session.getInstance(prop,new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(from, password);
			}
		});
		
		return session;
	}
}