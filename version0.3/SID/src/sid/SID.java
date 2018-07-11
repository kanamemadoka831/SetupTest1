package sid;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.security.MessageDigest;
import java.sql.*;
import java.util.Base64;

import javax.crypto.Cipher;
import javax.crypto.spec.SecretKeySpec;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.text.*;


public class SID {
	static int rows;
	static final String JDBC_DRIVER ="com.mysql.jdbc.Driver";  
	static final String DB_URL = "jdbc:mysql://localhost:3306/JiCheng?useSSL=false&serverTimezone = GMT";
	static final String USER = "root";
	static final String PASS = "Kanamemadoka_831";
	static public void setUpMysql()
	{
		Connection conn = null;
		Statement stmt = null;
		try{
			// 注册 JDBC 驱动
			Class.forName("com.mysql.jdbc.Driver");

			// 打开链接
			//			System.out.println("连接数据库...");
			conn = DriverManager.getConnection(DB_URL,USER,PASS);

			// 执行查询
			//			System.out.println(" 实例化Statement对象...");
			stmt = conn.createStatement();
			if(stmt!=null)
			{
				//
				String sqlSetUpTable="CREATE TABLE IF NOT EXISTS BuyInformation("
						+ "company_id VARCHAR(100) not NULL,"
						+"time_table VARCHAR(100) not NULL,"
						+ "machine_id VARCHAR(100) not NULL)";
				stmt.executeUpdate(sqlSetUpTable);

			}
		}catch(SQLException se){
			// 处理 JDBC 错误
			se.printStackTrace();
		}catch(Exception e){
			// 处理 Class.forName 错误
			e.printStackTrace();
		}finally{
			// 关闭资源
			try{
				if(stmt!=null) stmt.close();
			}catch(SQLException se2){
			}// 什么都不做
			try{
				if(conn!=null) conn.close();
			}catch(SQLException se){
				se.printStackTrace();
			}
		}
		//		System.out.println("Goodbye!");
	}
	public static String encryptDES(String encryptString, String encryptKey) throws Exception {
		int length=encryptString.getBytes("unicode").length;
		StringBuffer sb=new StringBuffer();
		sb.append(encryptString);
		if(length<8){
			for(int i=8-length;i<8;i++){
				sb.append("1");
			}
			encryptString=sb.toString();
			//			System.out.println(encryptString);
		}
		else if(length==8){
		}
		else {
			encryptString=encryptString.substring(0, 8);
			//			System.out.println(encryptString);
			//			System.out.println(encryptString.getBytes().length);
		}
		SecretKeySpec key = new SecretKeySpec(encryptKey.getBytes(), "DES");
		Cipher cipher = Cipher.getInstance("DES/ECB/NoPadding");
		cipher.init(Cipher.ENCRYPT_MODE, key);
		byte[] encryptedData = cipher.doFinal(encryptString.getBytes());
		return Base64.getEncoder().encodeToString(encryptedData);
	}
	public static String decryptDES(String decryptString, String decryptKey) throws Exception {
		//		int length=decryptKey.getBytes("unicode").length;
		//		StringBuffer sb=new StringBuffer();
		//		if(length<8){
		//			for(int i=8-length;i<8;i++){
		//				sb.append("1");
		//			}
		//			decryptKey=sb.toString();
		//			System.out.println(decryptKey);
		//		}
		//		else if(length==8){
		//		}
		//		else {
		//			decryptKey=decryptKey.substring(2, 10);
		//			System.out.println(decryptKey);
		//			System.out.println(decryptKey.getBytes().length);
		//		}
		byte[] byteMi =  Base64.getDecoder().decode(decryptString);
		SecretKeySpec key = new SecretKeySpec(decryptKey.getBytes(), "DES");
		Cipher cipher = Cipher.getInstance("DES/ECB/NoPadding");
		cipher.init(Cipher.DECRYPT_MODE, key);
		byte decryptedData[] = cipher.doFinal(byteMi);
		return new String(decryptedData);
	}
	private static String MD5(String s) {
		try {
			MessageDigest md = MessageDigest.getInstance("MD5");
			byte[] bytes = md.digest(s.getBytes("utf-8"));
			return toHex(bytes);
		}
		catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	private static String toHex(byte[] bytes) {

		final char[] HEX_DIGITS = "0123456789ABCDEF".toCharArray();
		StringBuilder ret = new StringBuilder(bytes.length * 2);
		for (int i=0; i<bytes.length; i++) {
			ret.append(HEX_DIGITS[(bytes[i] >> 4) & 0x0f]);
			ret.append(HEX_DIGITS[bytes[i] & 0x0f]);
		}
		return ret.toString();
	}
	public static void main(String[] args)
	{
		JFrame frame = new JFrame("Login Example");
		frame.setSize(350, 200);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		JPanel panel = new JPanel();    
		frame.add(panel);
		panel.setLayout(null);
		JLabel userLabel = new JLabel("购买方公司名:");
		userLabel.setBounds(10,20,100,25);
		panel.add(userLabel);
		JTextField userText = new JTextField(20);
		userText.setBounds(130,20,165,25);
		panel.add(userText);
		JLabel machineLabel = new JLabel("购买方机器码：");
		machineLabel.setBounds(10,50,100,25);
		panel.add(machineLabel);
		JTextField machineText = new JTextField(20);
		machineText.setBounds(130,50,165,25);
		panel.add(machineText);
		JButton loginButton = new JButton("生成注册码");
		loginButton.setBounds(10, 80, 80, 25);
		panel.add(loginButton);
		loginButton.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				int numbers;
				String oriCompuCodes=machineText.getText();
				String[] compuCodes=oriCompuCodes.split(";");
				numbers=compuCodes.length;
				rows=numbers;
				//				for(String num:compuCodes)
				//				{
				//					System.out.println(num);
				//				}
				CustomUUID uuid=new CustomUUID(17,5);
				setUpMysql();
				Long[] arr=new Long[numbers];
				String JDBC_DRIVER ="com.mysql.jdbc.Driver";  
				String DB_URL = "jdbc:mysql://localhost:3306/JiCheng?useSSL=false&serverTimezone = GMT";
				String USER = "root";
				String PASS = "Kanamemadoka_831";
				String companyName=userText.getText();
				Connection conn = null;
				Statement stmt = null;
				try{
					// 注册 JDBC 驱动
					Class.forName("com.mysql.jdbc.Driver");

					// 打开链接
					//					System.out.println("连接数据库...");
					conn = DriverManager.getConnection(DB_URL,USER,PASS);

					// 执行查询
					//					System.out.println(" 实例化Statement对象...");
					stmt = conn.createStatement();
					if(stmt!=null)
					{
						//
						for(int i=0;i<numbers;i++)
						{
							arr[i]=uuid.generate();
							//							System.out.println(arr[i].toString().substring(2, 10));
							//							System.out.println("密文的长度为"+arr[i].toString().substring(2, 10).getBytes().length);
							String finalCode=encryptDES(compuCodes[i],arr[i].toString().substring(2, 10))+MD5(compuCodes[i]+arr[i].toString());
							String encode1=encryptDES(compuCodes[i],arr[i].toString().substring(2, 10));
							//							System.out.println(compuCodes[i]);
							//							System.out.println(arr[i].toString());
							//							System.out.println(encode1);
							//							System.out.println(decryptDES(encode1,arr[i].toString().substring(2, 10)));
							String  sql = "INSERT INTO buyinformation " +
									"VALUES ('"+companyName+"', '"+arr[i]+"', '"+compuCodes[i]+"')";
							int result = stmt.executeUpdate(sql);
							//此处生成注册码或者注册文件
							String filePath = this.getClass().getProtectionDomain().getCodeSource().getLocation().getPath();
							filePath=filePath+(i+1)+".txt";
							System.out.println(filePath);
							File storeFile=new File(filePath);
							if(storeFile.exists()) {

								FileOutputStream fileStream;
								try {
									fileStream = new FileOutputStream(storeFile,true);
									fileStream.write(finalCode.getBytes());
									fileStream.close();
								} catch (FileNotFoundException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								} catch (IOException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
							}
							else {
								try {
									storeFile.createNewFile();

									FileOutputStream fileStream;
									try {
										fileStream = new FileOutputStream(storeFile);
										fileStream.write(finalCode.getBytes());
										fileStream.close();
									} catch (FileNotFoundException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									} catch (IOException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}	
								} catch (IOException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();

								}

							}
						}

					}
				}catch(SQLException se){
					// 处理 JDBC 错误
					se.printStackTrace();
				}catch(Exception e){
					// 处理 Class.forName 错误
					e.printStackTrace();
				}finally{
					// 关闭资源
					try{
						if(stmt!=null) stmt.close();
					}catch(SQLException se2){
					}// 什么都不做
					try{
						if(conn!=null) conn.close();
					}catch(SQLException se){
						se.printStackTrace();
					}
				}
			}

		});

		frame.setVisible(true);
	}
}

//	public static void main(String[] args)
//	{
//		CustomUUID uuid=new CustomUUID(5,5);
//		setUpMysql();
//		int length=Integer.valueOf(args[0]);
//		Long[] arr=new Long[length];
//		String JDBC_DRIVER ="com.mysql.jdbc.Driver";  
//		String DB_URL = "jdbc:mysql://localhost:3306/JiCheng?useSSL=false&serverTimezone = GMT";
//		String USER = "root";
//		String PASS = "Kanamemadoka_831";
//		String companyName="xxCompany";
//		String[] machineID=new String[length]; 
//		Connection conn = null;
//		Statement stmt = null;
//		for(int i=0;i<length;i++) {
//			machineID[i]="sdsdsdd45648"+i+"s4da56a16s4d856sa16";
//		}
//		try{
//			// 注册 JDBC 驱动
//			Class.forName("com.mysql.jdbc.Driver");
//
//			// 打开链接
//			System.out.println("连接数据库...");
//			conn = DriverManager.getConnection(DB_URL,USER,PASS);
//
//			// 执行查询
//			System.out.println(" 实例化Statement对象...");
//			stmt = conn.createStatement();
//			if(stmt!=null)
//			{
//				//
//				for(int i=0;i<Integer.valueOf(length);i++)
//				{
//					arr[i]=uuid.generate();
//					String  sql = "INSERT INTO buyinformation " +
//			                   "VALUES ('"+companyName+"', '"+arr[i]+"', '"+machineID[i]+"')";
//
//					
//					int result = stmt.executeUpdate(sql);
//					
//				}
//
//			}
//		}catch(SQLException se){
//			// 处理 JDBC 错误
//			se.printStackTrace();
//		}catch(Exception e){
//			// 处理 Class.forName 错误
//			e.printStackTrace();
//		}finally{
//			// 关闭资源
//			try{
//				if(stmt!=null) stmt.close();
//			}catch(SQLException se2){
//			}// 什么都不做
//			try{
//				if(conn!=null) conn.close();
//			}catch(SQLException se){
//				se.printStackTrace();
//			}
//		}
//
//
//
//		//		String filePath="D:\\JavaWorkspace\\SID\\saved.txt";
//		//		File storeFile=new File(filePath);
//		//		if(storeFile.exists()) {
//		//			for(int i=0;i<Integer.valueOf(args[0]);i++)
//		//			{
//		//				FileOutputStream fileStream;
//		//				try {
//		//					fileStream = new FileOutputStream(storeFile,true);
//		//					Long time=uuid.generate();
//		//					String timeInput=time.toString()+"\n";
//		//					fileStream.write(timeInput.getBytes());
//		//				} catch (FileNotFoundException e) {
//		//					// TODO Auto-generated catch block
//		//					e.printStackTrace();
//		//				} catch (IOException e) {
//		//					// TODO Auto-generated catch block
//		//					e.printStackTrace();
//		//				}
//		//				
//		//				
//		//			}
//		//		}
//		//		else {
//		//			try {
//		//				storeFile.createNewFile();
//		//				for(int i=0;i<Integer.valueOf(args[0]);i++)
//		//				{
//		//					FileOutputStream fileStream;
//		//					try {
//		//						fileStream = new FileOutputStream(storeFile);
//		//						fileStream.write((int) uuid.generate());
//		//					} catch (FileNotFoundException e) {
//		//						// TODO Auto-generated catch block
//		//						e.printStackTrace();
//		//					} catch (IOException e) {
//		//						// TODO Auto-generated catch block
//		//						e.printStackTrace();
//		//					}	
//		//				}
//		//			} catch (IOException e) {
//		//				// TODO Auto-generated catch block
//		//				e.printStackTrace();
//		//				
//		//			}
//		//			
//		//		}
//	}
//}

