package org.vishal.aparupdater;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.googleapis.util.Utils;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

import AccountDetail.MainClass;
import Environment.Bartos;
import Environment.EMS;
import Environment.HDGrant;
import Environment.MPNCompany;
import Environment.MPNServices;
import Environment.MWG;
import Environment.Pennergy;
import Utils.DateGenerator;


public class SheetUpdater {
	private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String TOKENS_DIRECTORY_PATH = "tokens";
	private static DateGenerator date;

	/**
	 * Global instance of the scopes required by this quickstart.
	 * If modifying these scopes, delete your previously saved tokens/ folder.
	 */
	private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
	private static final String CREDENTIALS_FILE_PATH = File.separator+"credentials.json";

	/**
	 * Creates an authorized Credential object.
	 * @param HTTP_TRANSPORT The network HTTP Transport.
	 * @return An authorized Credential object.
	 * @throws IOException If the credentials.json file cannot be found.
	 */
	private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
		// Load client secrets.
		InputStream in = SheetUpdater.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
		if (in == null) {
			throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
		}
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
				HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
				.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
				.setAccessType("offline")
				.build();
		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
		return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
	}

	public static void main(String[] args) throws IOException, GeneralSecurityException {
		// Build a new authorized API client service.
		MainClass.main(args);
		date = new DateGenerator();
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		final String spreadsheetId = "1k7BNjjkekwhvVzbwXA1kilGEWCIczxNCGHnaQn0CqqY";
		Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
				.setApplicationName(APPLICATION_NAME)
				.build();

		SheetUpdater sheetUpdate = new SheetUpdater();
//		sheetUpdate.updateBartos(service, spreadsheetId);
//		sheetUpdate.updateHDGrant(service, spreadsheetId);
//		sheetUpdate.updateMWG(service, spreadsheetId);
//		sheetUpdate.updateMPNCompany(service, spreadsheetId);
//		sheetUpdate.updateMPNServices(service, spreadsheetId);
//		sheetUpdate.updatePennergy(service, spreadsheetId);
		sheetUpdate.updateEMS(service, spreadsheetId);
	}



	/**
	 * Updates the AP and AR value of Bartos Instance.
	 * @param service The object of spreadsheet service.
	 * @param spreadsheetId The ID of the Bartos spreadsheet.
	 */
	@SuppressWarnings("unchecked")
	private void updateBartos(Sheets service, String spreadsheetId) {

		String range = "Bartos";
		Bartos bartos = new Bartos();
		BigDecimal ar = bartos.getAR();
		BigDecimal ap = bartos.getAP();
		BigDecimal inv = BigDecimal.valueOf(bartos.getInvValue());
		BigDecimal ar_gl = bartos.get_GLAccountBalances( bartos.getPeriodandYear().get("period"), bartos.getPeriodandYear().get("year") ).get("Accounts Receivable");
		BigDecimal inv_gl = bartos.get_GLAccountBalances( bartos.getPeriodandYear().get("period"), bartos.getPeriodandYear().get("year") ).get("Inventory");
		BigDecimal ap_gl = bartos.get_GLAccountBalances( bartos.getPeriodandYear().get("period"), bartos.getPeriodandYear().get("year") ).get("Accounts Payable");


		List<List<Object>> values = Arrays.asList(
				Arrays.asList(
						// Cell values ...
						(Object)date.getCurrentDate(), (Object)ar, (Object)ar_gl, (Object)ar_gl.subtract(ar), (Object)ap, (Object)ap_gl, (Object)ap_gl.subtract(ap),
						(Object)inv,
						(Object)inv_gl,
						(Object)inv.subtract(inv_gl)
						)
				// Additional rows ...
				);
		ValueRange body = new ValueRange()
				.setValues(values);

		try {
			service.spreadsheets().values().append(spreadsheetId, range, body)
			.setValueInputOption("USER_ENTERED")
			.execute();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.printf("\nBARTOS instance updated \n");
	}

	/**
	 * Updates the AP and AR value of HDGrant Instance.
	 * @param service The object of spreadsheet service.
	 * @param spreadsheetId The ID of the HDGrant spreadsheet.
	 */
	@SuppressWarnings("unchecked")

	private void updateHDGrant(Sheets service, String spreadsheetId) {

		String range = "HD Grant";
		HDGrant hdgrant = new HDGrant();
		BigDecimal ar = hdgrant.getAR();
		BigDecimal ap = hdgrant.getAP();
		BigDecimal inv = BigDecimal.valueOf(hdgrant.getInvValue());
		BigDecimal ar_gl = hdgrant.get_GLAccountBalances( hdgrant.getPeriodandYear().get("period"), hdgrant.getPeriodandYear().get("year") ).get("Accounts Receivable");
		BigDecimal inv_gl = hdgrant.get_GLAccountBalances( hdgrant.getPeriodandYear().get("period"), hdgrant.getPeriodandYear().get("year") ).get("Inventory");
		BigDecimal ap_gl = hdgrant.get_GLAccountBalances( hdgrant.getPeriodandYear().get("period"), hdgrant.getPeriodandYear().get("year") ).get("Accounts Payable");


		List<List<Object>> values = Arrays.asList(
				Arrays.asList(
						// Cell values ...
						(Object)date.getCurrentDate(), (Object)ar, (Object)ar_gl, (Object)ar_gl.subtract(ar), (Object)ap, (Object)ap_gl, (Object)ap_gl.subtract(ap),
						(Object)inv,
						(Object)inv_gl,
						(Object)inv.subtract(inv_gl)
						)
				// Additional rows ...
				);
		ValueRange body = new ValueRange()
				.setValues(values);

		try {
			service.spreadsheets().values().append(spreadsheetId, range, body)
			.setValueInputOption("USER_ENTERED")
			.execute();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.printf("HDGrant instance updated \n");
	}

	/**
	 * Updates the AP and AR value of MWG Instance.
	 * @param service The object of spreadsheet service.
	 * @param spreadsheetId The ID of the MWG spreadsheet.
	 */
	@SuppressWarnings("unchecked")
	private void updateMWG(Sheets service, String spreadsheetId) {

		String range = "MWG";
		MWG mwg = new MWG();
		BigDecimal ar = mwg.getAR();
		BigDecimal ap = mwg.getAP();
		BigDecimal inv = BigDecimal.valueOf(mwg.getInvValue());
		BigDecimal ar_gl = mwg.get_GLAccountBalances( mwg.getPeriodandYear().get("period"), mwg.getPeriodandYear().get("year") ).get("Accounts Receivable");
		BigDecimal inv_gl = mwg.get_GLAccountBalances( mwg.getPeriodandYear().get("period"), mwg.getPeriodandYear().get("year") ).get("Inventory");
		BigDecimal ap_gl = mwg.get_GLAccountBalances( mwg.getPeriodandYear().get("period"), mwg.getPeriodandYear().get("year") ).get("Accounts Payable");


		List<List<Object>> values = Arrays.asList(
				Arrays.asList(
						// Cell values ...
						(Object)date.getCurrentDate(), (Object)ar, (Object)ar_gl, (Object)ar_gl.subtract(ar), (Object)ap, (Object)ap_gl, (Object)ap_gl.subtract(ap),
						(Object)inv,
						(Object)inv_gl,
						(Object)inv.subtract(inv_gl)
						)
				// Additional rows ...
				);
		ValueRange body = new ValueRange()
				.setValues(values);

		try {
			service.spreadsheets().values().append(spreadsheetId, range, body)
			.setValueInputOption("USER_ENTERED")
			.execute();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.printf("MWG instance updated \n");
	}

	/**
	 * Updates the AP and AR value of MPNServices Instance.
	 * @param service The object of spreadsheet service.
	 * @param spreadsheetId The ID of the MPNServices spreadsheet.
	 */
	@SuppressWarnings("unchecked")
	private void updateMPNServices(Sheets service, String spreadsheetId) {

		String range = "MPN Services";
		MPNServices mpns = new MPNServices();
		BigDecimal ar = mpns.getAR();
		BigDecimal ap = mpns.getAP();
		BigDecimal inv = BigDecimal.valueOf(mpns.getInvValue());
		BigDecimal ar_gl = mpns.get_GLAccountBalances( mpns.getPeriodandYear().get("period"), mpns.getPeriodandYear().get("year") ).get("Accounts Receivable");
		BigDecimal inv_gl = mpns.get_GLAccountBalances( mpns.getPeriodandYear().get("period"), mpns.getPeriodandYear().get("year") ).get("Inventory");
		BigDecimal ap_gl = mpns.get_GLAccountBalances( mpns.getPeriodandYear().get("period"), mpns.getPeriodandYear().get("year") ).get("Accounts Payable");


		List<List<Object>> values = Arrays.asList(
				Arrays.asList(
						// Cell values ...
						(Object)date.getCurrentDate(), (Object)ar, (Object)ar_gl, (Object)ar_gl.subtract(ar), (Object)ap, (Object)ap_gl, (Object)ap_gl.subtract(ap),
						(Object)inv,
						(Object)inv_gl,
						(Object)inv.subtract(inv_gl)
						)
				// Additional rows ...
				);
		ValueRange body = new ValueRange()
				.setValues(values);

		try {
			service.spreadsheets().values().append(spreadsheetId, range, body)
			.setValueInputOption("USER_ENTERED")
			.execute();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.printf("MPN Sertices instance updated \n");
	}

	/**
	 * Updates the AP and AR value of MPNCompany Instance.
	 * @param service The object of spreadsheet service.
	 * @param spreadsheetId The ID of the MPNCompany spreadsheet.
	 */
	@SuppressWarnings("unchecked")
	private void updateMPNCompany(Sheets service, String spreadsheetId) {

		String range = "MPN Company";
		MPNCompany mpnc = new MPNCompany();
		BigDecimal ar = mpnc.getAR();
		BigDecimal ap = mpnc.getAP();
		BigDecimal inv = BigDecimal.valueOf(mpnc.getInvValue());
		BigDecimal ar_gl = mpnc.get_GLAccountBalances( mpnc.getPeriodandYear().get("period"), mpnc.getPeriodandYear().get("year") ).get("Accounts Receivable");
		BigDecimal inv_gl = mpnc.get_GLAccountBalances( mpnc.getPeriodandYear().get("period"), mpnc.getPeriodandYear().get("year") ).get("Inventory");
		BigDecimal ap_gl = mpnc.get_GLAccountBalances( mpnc.getPeriodandYear().get("period"), mpnc.getPeriodandYear().get("year") ).get("Accounts Payable");


		List<List<Object>> values = Arrays.asList(
				Arrays.asList(
						// Cell values ...
						(Object)date.getCurrentDate(), (Object)ar, (Object)ar_gl, (Object)ar_gl.subtract(ar), (Object)ap, (Object)ap_gl, (Object)ap_gl.subtract(ap),
						(Object)inv,
						(Object)inv_gl,
						(Object)inv.subtract(inv_gl)
						)
				// Additional rows ...
				);
		ValueRange body = new ValueRange()
				.setValues(values);

		try {
			service.spreadsheets().values().append(spreadsheetId, range, body)
			.setValueInputOption("USER_ENTERED")
			.execute();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.printf("MPN Company instance updated \n");
	}

	/**
	 * Updates the AP and AR value of Pennergy Instance.
	 * @param service The object of spreadsheet service.
	 * @param spreadsheetId The ID of the Pennergy spreadsheet.
	 */
	@SuppressWarnings("unchecked")
	private void updatePennergy(Sheets service, String spreadsheetId) {

		String range = "Pennergy";
		Pennergy pennergy = new Pennergy();
		BigDecimal ar = pennergy.getAR();
		BigDecimal ap = pennergy.getAP();
		BigDecimal inv = BigDecimal.valueOf(pennergy.getInvValue());
		BigDecimal ar_gl = pennergy.get_GLAccountBalances( pennergy.getPeriodandYear().get("period"), pennergy.getPeriodandYear().get("year") ).get("Accounts Receivable");
		BigDecimal inv_gl = pennergy.get_GLAccountBalances( pennergy.getPeriodandYear().get("period"), pennergy.getPeriodandYear().get("year") ).get("Inventory");
		BigDecimal ap_gl = pennergy.get_GLAccountBalances( pennergy.getPeriodandYear().get("period"), pennergy.getPeriodandYear().get("year") ).get("Accounts Payable");


		List<List<Object>> values = Arrays.asList(
				Arrays.asList(
						// Cell values ...
						(Object)date.getCurrentDate(), (Object)ar, (Object)ar_gl, (Object)ar_gl.subtract(ar), (Object)ap, (Object)ap_gl, (Object)ap_gl.subtract(ap),
						(Object)inv,
						(Object)inv_gl,
						(Object)inv.subtract(inv_gl)
						)
				// Additional rows ...
				);
		ValueRange body = new ValueRange()
				.setValues(values);

		try {
			service.spreadsheets().values().append(spreadsheetId, range, body)
			.setValueInputOption("USER_ENTERED")
			.execute();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.printf("Pennergy instance updated \n");
	}
	
	/**
	 * Updates the AP and AR value of EMS Instance.
	 * @param service The object of spreadsheet service.
	 * @param spreadsheetId The ID of the EMS spreadsheet.
	 */
	@SuppressWarnings("unchecked")
	private void updateEMS(Sheets service, String spreadsheetId) {

		String range = "EMS";
		EMS ems = new EMS();
		BigDecimal ar = ems.getAR();
		BigDecimal ap = ems.getAP();
		BigDecimal inv = BigDecimal.valueOf(ems.getInvValue());
		BigDecimal ar_gl = ems.get_GLAccountBalances( ems.getPeriodandYear().get("period"), ems.getPeriodandYear().get("year") ).get("Accounts Receivable");
		BigDecimal inv_gl = ems.get_GLAccountBalances( ems.getPeriodandYear().get("period"), ems.getPeriodandYear().get("year") ).get("Inventory");
		BigDecimal ap_gl = ems.get_GLAccountBalances( ems.getPeriodandYear().get("period"), ems.getPeriodandYear().get("year") ).get("Accounts Payable");


		List<List<Object>> values = Arrays.asList(
				Arrays.asList(
						// Cell values ...
						(Object)date.getCurrentDate(), (Object)ar, (Object)ar_gl, (Object)ar.subtract(ar_gl), (Object)ap, (Object)ap_gl, (Object)ap.subtract(ap_gl),
						(Object)inv,
						(Object)inv_gl,
						(Object)inv.subtract(inv_gl)
						)
				// Additional rows ...
				);
		ValueRange body = new ValueRange()
				.setValues(values);

		try {
			service.spreadsheets().values().append(spreadsheetId, range, body)
			.setValueInputOption("USER_ENTERED")
			.execute();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.printf("EMS instance updated \n");
	}
}


