package com.mindstorm.sfexport;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;

import javax.servlet.http.*;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

@SuppressWarnings("serial")
public class ExportSFContentServlet extends HttpServlet {

   /*******************************************************************************************************
   The HTTP Get request is going to be invoked from Sales Force Custom Link in Opportunity form as follows:
   http://exportsfdata.cloudfoundry.com/export?oppname={!Opportunity.Name}&
                                               oppamount={!Opportunity.Amount}&
											   oppstage={!Opportunity.StageName}&
											   accountname={!Opportunity.Account}&
											   currentusername={!$User.FirstName}&
											   accountbillingaddress={!Account.BillingAddress}
   ********************************************************************************************************/

	public void doGet(HttpServletRequest request, HttpServletResponse response) throws IOException {
	
	    /****************************************************************************************
		Currently the Servlet processes the following parameters. You can change them in the code
		in case you plan to extract out other variables.
		---------------------------------------------------------------------------------
		| Parameter Name	     |      Description                                     |
		---------------------------------------------------------------------------------
		| oppname	                This is the opportunity name                        |
		| oppamount	                This is the amount in $$$                           |
		| oppstage	                This is the opportunity stage                       |
		| accountname	            This is the account name                            |
		| currentusername	        The current user name of the account logged in      |
		| accountbillingaddress	    The account billing address                         |
		---------------------------------------------------------------------------------
		*****************************************************************************************/

		WordprocessingMLPackage wordMLPackage;

		try {
		
			wordMLPackage = WordprocessingMLPackage.createPackage();
			MainDocumentPart mainPart = wordMLPackage.getMainDocumentPart();

			// create some styled heading and add the text values
			// The first parameter is the Word Styles, the second parameter is the content
			mainPart.addStyledParagraphOfText("Title", "Salesforce Opportunity Details");
			mainPart.addStyledParagraphOfText("Subtitle", "Generated at " + Calendar.getInstance().getTime().toString());
			mainPart.addStyledParagraphOfText("Subtle Emphasis","Account Name : " + request.getParameter("accountname"));
			mainPart.addStyledParagraphOfText("Subtle Emphasis","Billing Address : " + request.getParameter("accountbillingaddress"));
			mainPart.addStyledParagraphOfText("Subtle Emphasis","Opportunity Name: " + request.getParameter("oppname"));
			mainPart.addStyledParagraphOfText("Subtle Emphasis","Opportunity Amount : " + request.getParameter("oppamount"));
			mainPart.addStyledParagraphOfText("Subtle Emphasis","Opportunity Stage : " + request.getParameter("oppstage"));
			mainPart.addStyledParagraphOfText("Subtle Emphasis","User Name : " + request.getParameter("currentusername"));
			
			//Save the contents in a temporary file
			File file = File.createTempFile("wordexport-", ".docx");
			wordMLPackage.save(file);
			
			//Now send back the response to the Server
			response.setHeader("Content-disposition", "attachment; filename=opportunitydetails.docx");
			response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
			OutputStream out = response.getOutputStream();
			FileInputStream in = new FileInputStream(file);
			byte[] buffer = new byte[4096];
			int length;
			while ((length = in.read(buffer)) > 0){
			    out.write(buffer, 0, length);
			}
			in.close();
			out.flush();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		}
	}
}
