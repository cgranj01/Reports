
package principal;

import java.awt.print.PrinterException;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.MalformedURLException;
import java.net.URL;
import java.security.GeneralSecurityException;
import java.security.InvalidAlgorithmParameterException;
import java.security.InvalidKeyException;
import java.security.NoSuchAlgorithmException;
import java.security.spec.InvalidKeySpecException;
import java.text.SimpleDateFormat;
import java.util.GregorianCalendar;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.crypto.BadPaddingException;
import javax.crypto.Cipher;
import javax.crypto.IllegalBlockSizeException;
import javax.crypto.NoSuchPaddingException;
import javax.crypto.SecretKey;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.PBEKeySpec;
import javax.crypto.spec.PBEParameterSpec;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.commons.configuration.ConfigurationException;
import org.apache.commons.configuration.PropertiesConfiguration;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.pdfbox.pdmodel.PDDocument;

import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

public class Reports
{
	private PropertiesConfiguration	prop;
	private File					f;
	private static final char []	PASSWORD	= "vffd4njomm2rkghr6oeslhys88m".toCharArray ();
	private static final byte []	SALT		= {(byte) 0xdc, (byte) 0x3F, (byte) 0x10, (byte) 0x19, (byte) 0xbe,
			(byte) 0x33, (byte) 0x1A, (byte) 0xA9,};
	
	private static Logger logger = LogManager.getLogger(Reports.class);
	
	public Reports ()
	{
		try
		{
			logger.info ("Demarrage traitement...");
			this.prop = new PropertiesConfiguration ("config.properties");
			this.prop.setAutoSave (true);
		}
		catch (ConfigurationException aCe)
		{
			logger.error(aCe.getMessage ());
		}
		
		String aURL = this.prop.getString ("url");
		String [] aNomFichier= this.prop.getString ("pathname").split("\\.");
		String aNombreCopie = this.prop.getString ("nbcopyprinter");
		String [] aNomFeuille = this.prop.getStringArray ("feuilles");
		String [] aDestMail = this.prop.getStringArray ("to");
		
		GregorianCalendar cal = new GregorianCalendar ();
		cal.add (GregorianCalendar.DATE, - 1);
		SimpleDateFormat sdf = new SimpleDateFormat ("yyyy_MM");
		String jour = sdf.format (cal.getTime ());
		String fichier = aNomFichier[0] + "_"+ jour + "." + aNomFichier[1];
		this.f = new File (fichier);

		if ( ! aURL.equalsIgnoreCase (""))
		{
			try
			{
				URL url = new URL (aURL);
				FileUtils.copyURLToFile(url, this.f);
				logger.info ("fichier {} ecrit correctement", fichier);
			}
			catch (MalformedURLException e)
			{
				logger.error(e.getMessage ());
			}
			catch (IOException e)
			{
				logger.error("Erreur acces fichier : {}", e.getMessage ());
			}
		}
		
		
		if (Integer.parseInt (aNombreCopie) > 0)
		{
			if (aNomFichier[1].equalsIgnoreCase ("xls") || aNomFichier[1].equalsIgnoreCase ("xlsx"))
			{
				for (int i = 0; i < aNomFeuille.length; i ++ )
				{
					this.printFileXLS (this.f, aNombreCopie, aNomFeuille [i]);
					try
					{
						Thread.sleep (1000L);
					}
					catch (InterruptedException e)
					{
						e.printStackTrace ();
					}
				}
			}
				
			if (aNomFichier[1].equalsIgnoreCase ("pdf"))
			{
				this.printFilePDF (this.f, aNombreCopie);
			}
		}
				
		if (aDestMail.length > 0 )
		{
			this.sendEmail ();
		}
	}

	private void printFilePDF (File afile, String aCopyNumber)
	{
		try
		{
			logger.info ("impression demandée pour {} en {} exemplaire", afile, aCopyNumber);
			PDDocument doc = PDDocument.load (afile);
			doc.silentPrint ();
			logger.info ("impression envoyée à l'imprimante par défaut du systeme");
		}
		catch (PrinterException e)
		{
			logger.error("Erreur impression : {}", e.getMessage ());
		}
		catch (IOException e)
		{
			logger.error("Erreur acces fichier pour impression : {}", e.getMessage ());
		}
	}

	private void printFileXLS (File afile, String aCopyNumber, String aName)
	{
		try
		{
			String vbs = "Dim AppExcel\r\n" + "Set AppExcel = CreateObject(\"Excel.application\")\r\n"
					+ "AppExcel.Workbooks.Open(\"" + afile + "\")\r\n" + "AppExcel.Sheets(\"" + aName
					+ "\").PrintOut ,," + aCopyNumber + "\r\n" + "AppExcel.Workbooks.Close\r\n"
					// + "appExcel.ActiveWindow.SelectedSheets.PrintOut\r\n"
					+ "Appexcel.Quit\r\n" + "Set AppExcel = Nothing";
			File vbScript = File.createTempFile ("vbScript", ".vbs");
			System.out.println (vbs);
			FileWriter fw = new java.io.FileWriter (vbScript);
			fw.write (vbs);
			fw.close ();
			Runtime.getRuntime ().exec ("cmd /c" + vbScript.getPath ());
			// vbScript.deleteOnExit();
		}
		catch (Exception e)
		{
			logger.error(e.getMessage ());
		}
	}

	private String encrypt (String property) throws GeneralSecurityException, UnsupportedEncodingException
	{
		SecretKeyFactory keyFactory = SecretKeyFactory.getInstance ("PBEWithMD5AndDES");
		SecretKey key = keyFactory.generateSecret (new PBEKeySpec (PASSWORD));
		Cipher pbeCipher = Cipher.getInstance ("PBEWithMD5AndDES");
		pbeCipher.init (Cipher.ENCRYPT_MODE, key, new PBEParameterSpec (SALT, 20));
		return base64Encode (pbeCipher.doFinal (property.getBytes ("UTF-8")));
	}

	private String base64Encode (byte [] bytes)
	{
		// NB: This class is internal, and you probably should use another impl
		return new BASE64Encoder ().encode (bytes);
	}

	private String decrypt (String property)
	{
		SecretKeyFactory keyFactory;
		try
		{
			keyFactory = SecretKeyFactory.getInstance ("PBEWithMD5AndDES");
			SecretKey key = keyFactory.generateSecret (new PBEKeySpec (PASSWORD));
			Cipher pbeCipher = Cipher.getInstance ("PBEWithMD5AndDES");
			pbeCipher.init (Cipher.DECRYPT_MODE, key, new PBEParameterSpec (SALT, 20));
			String string = new String (pbeCipher.doFinal (base64Decode (property)), "UTF-8");
			return string;
		}
		catch (NoSuchAlgorithmException e)
		{
			logger.error(e.getMessage ());
			return "";
		}
		catch (InvalidKeySpecException e)
		{
			logger.error(e.getMessage ());
			return "";
		}
		catch (NoSuchPaddingException e)
		{
			logger.error(e.getMessage ());
			return "";
		}
		catch (InvalidKeyException e)
		{
			logger.error(e.getMessage ());
			return "";
		}
		catch (InvalidAlgorithmParameterException e)
		{
			logger.error(e.getMessage ());
			return "";
		}
		catch (UnsupportedEncodingException e)
		{
			logger.error(e.getMessage ());
			return "";
		}
		catch (IllegalBlockSizeException e)
		{
			logger.error(e.getMessage ());
			return "";
		}
		catch (BadPaddingException e)
		{
			logger.error(e.getMessage ());
			return "";
		}
		catch (IOException e)
		{
			logger.error(e.getMessage ());
			return "";
		}

	}

	private byte [] base64Decode (String property) throws IOException
	{
		return new BASE64Decoder ().decodeBuffer (property);
	}

	public void Encryption (String aKey, String aValue, String aKeytoclear)
	{
		if ( ! aValue.equals (""))
		{
			String enc;
			try
			{
				enc = this.encrypt (aValue);
				this.prop.setProperty (aKey, enc);
				this.prop.setProperty (aKeytoclear, "");
			}
			catch (UnsupportedEncodingException e)
			{
				logger.error(e.getMessage ());
			}
			catch (GeneralSecurityException e)
			{
				logger.error(e.getMessage ());
			}

		}

	}

	public void sendEmail ()
	{
		GregorianCalendar cal = new GregorianCalendar ();
		cal.add (GregorianCalendar.DATE, - 1);
		SimpleDateFormat sdf = new SimpleDateFormat ("EEEE dd MMM yyyy");
		String jour = sdf.format (cal.getTime ());

		String aSenders = this.prop.getString ("from");
		String [] aReceiver = this.prop.getStringArray ("to");
		String aHost = this.prop.getString ("hote_smtp");
		String aPort = this.prop.getString ("port");
		String aAuth = this.prop.getString ("authentification");
		String aTLS = this.prop.getString ("tls");
		String aSSL = this.prop.getString ("ssl");

		String passwordclair = (String) this.prop.getProperty ("accountpwd");
		String accountclair = (String) this.prop.getProperty ("accountname");

		this.Encryption ("accountpwdencrypt", passwordclair, "accountpwd");
		this.Encryption ("accountnameencrypt", accountclair, "accountname");
		String passwordcrypte = (String) this.prop.getProperty ("accountpwdencrypt");
		String accnamecrypte = (String) this.prop.getProperty ("accountnameencrypt");

		final String password = this.decrypt (passwordcrypte);
		final String username = this.decrypt (accnamecrypte);

		Properties props = new Properties ();
		props.put ("mail.smtp.auth", aAuth);
		props.put ("mail.smtp.starttls.enable", aTLS);
		props.put ("mail.smtp.host", aHost);
		props.put ("mail.smtp.port", aPort);
		props.put ("mail.smtp.ssl.enable", aSSL);

		// Get the Session object.
		Session session = Session.getInstance (props, new javax.mail.Authenticator ()
		{
			protected PasswordAuthentication getPasswordAuthentication ()
			{
				return new PasswordAuthentication (username, password);
			}
		});

		try
		{
			// Create a default MimeMessage object.
			Message message = new MimeMessage (session);

			// Set From: header field of the header.
			message.setFrom (new InternetAddress (aSenders));

			// Set To: header field of the header
			InternetAddress [] addressTo = new InternetAddress [aReceiver.length];
			for (int i = 0; i < aReceiver.length; i ++ )
			{
				addressTo [i] = new InternetAddress (aReceiver [i]);
			}
			message.setRecipients (Message.RecipientType.TO, addressTo);
			// message.setRecipients(Message.RecipientType.TO,
			// InternetAddress.parse(adressTo));

			// Set Subject: header field
			String aSubject = (String) this.prop.getProperty ("subject");
			message.setSubject (aSubject);

			// Create the message part
			String aBody = (String) this.prop.getProperty ("body");
			BodyPart messageBodyPart = new MimeBodyPart ();
			messageBodyPart.setContent (aBody + jour, "text/html");

			// Create a multipart message
			Multipart multipart = new MimeMultipart ();

			// Set text message part
			multipart.addBodyPart (messageBodyPart);

			// Part two is attachment
			messageBodyPart = new MimeBodyPart ();
			DataSource source = new FileDataSource (this.f);
			messageBodyPart.setDataHandler (new DataHandler (source));
			messageBodyPart.setFileName (this.f.getName ());
			multipart.addBodyPart (messageBodyPart);

			// Send the complete message parts
			message.setContent (multipart);

			// Send message
			Transport.send (message);

			logger.info ("email envoyé avec succes à : {}", StringUtils.join (aReceiver, ", "));

		}
		catch (MessagingException e)
		{
			System.out.println (e.getMessage ());
			logger.error ("erreur d'envoi : {}", e.getMessage ());
		}

	}
}
