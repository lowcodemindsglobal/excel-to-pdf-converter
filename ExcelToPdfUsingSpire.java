package excelservice;

import org.apache.log4j.Logger;
import com.appiancorp.suiteapi.common.Name;
import com.appiancorp.suiteapi.content.ContentConstants;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.content.exceptions.InvalidContentException;
import com.appiancorp.suiteapi.knowledge.Document;
import com.appiancorp.suiteapi.knowledge.DocumentDataType;
import com.appiancorp.suiteapi.knowledge.FolderDataType;
import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import com.appiancorp.suiteapi.process.framework.AppianSmartService;
import com.appiancorp.suiteapi.process.framework.Input;
import com.appiancorp.suiteapi.process.framework.MessageContainer;
import com.appiancorp.suiteapi.process.framework.Required;
import com.appiancorp.suiteapi.process.framework.SmartServiceContext;
import com.appiancorp.suiteapi.process.palette.PaletteInfo;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

@PaletteInfo(paletteCategory = "Appian Smart Services", palette = "Document Management")
public class ExcelToPdfUsingSpire extends AppianSmartService {

	private static final Logger LOG = Logger.getLogger(ExcelToPdfUsingSpire.class);
	@SuppressWarnings("unused")
	private final SmartServiceContext smartServiceCtx;
	private Long excelDocument;
	private String documentName;
	private Long saveInFolder;
	private Long newGeneratedDocument;
	private final ContentService contentService;
	Long createdDocument;
	String filePath;
	private final String extensionValue = "pdf";
	String xlFileName;
	String licenseFileName;
	public String licenseKey;
	private boolean errorOccured;
	private String errorMessage;

	public void run() throws SmartServiceException {

		try {

			xlFileName = contentService.getInternalFilename(excelDocument);

		} catch (InvalidContentException e1) {
			errorOccured = true;
			errorMessage = "InvalidContentException";
			LOG.error("InvalidContentException");
		}
		try {
			convertExcelToPdfUsingSpire(xlFileName, filePath, documentName);

		} catch (Exception e) {
			LOG.error("Exception : " + e);
		}
	}

	public ExcelToPdfUsingSpire(SmartServiceContext smartServiceCtx, ContentService cs_) {
		super();
		this.smartServiceCtx = smartServiceCtx;
		this.contentService = cs_;
	}

	private Long createDocument(String documentName, String extensionVal) {
		Document document;
		Long generatedDocument = null;
		try {
			document = new Document(saveInFolder, documentName, extensionVal);
			document.setState(ContentConstants.STATE_ACTIVE_PUBLISHED);
			document.setFileSystemId(ContentConstants.ALLOCATE_FSID);
			generatedDocument = contentService.create(document, ContentConstants.UNIQUE_FOR_PARENT);

			newGeneratedDocument = contentService.getVersion(generatedDocument, ContentConstants.VERSION_CURRENT)
					.getId();
		} catch (Exception e) {
			LOG.error("createDocument Exception : " + e);
		}

		return generatedDocument;

	}

	public void convertExcelToPdfUsingSpire(String excelDocumentPath, String outputPath, String documentName)
			throws Exception {

		Workbook workbook = null;
		Worksheet worksheet = null;
		try {
			if (!licenseKey.trim().isEmpty()) {

				com.spire.license.LicenseProvider.setLicenseKey(licenseKey);
			}
			workbook = new Workbook();
			workbook.loadFromFile(excelDocumentPath);

			for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
				worksheet = workbook.getWorksheets().get(i);
				// AutoFit column width and row height
				worksheet.getAllocatedRange().autoFitColumns();
				worksheet.getAllocatedRange().autoFitRows();

			}

			workbook.getConverterSetting().setSheetFitToPage(true);

		} catch (Exception e) {
			errorOccured = true;
			errorMessage = "Exception ERROR : " + e;

		} finally {

		}

		if (errorOccured == false) {
			createdDocument = createDocument(documentName, extensionValue);
			filePath = contentService.getInternalFilename(createdDocument);
			workbook.saveToFile(filePath, FileFormat.PDF);
			contentService.setSizeOfDocumentVersion(createdDocument);
		}

	}

	public void onSave(MessageContainer messages) {
	}

	public void validate(MessageContainer messages) {
	}

	@Input(required = Required.ALWAYS)
	@Name("excelDocument")
	@DocumentDataType
	public void setExcelDocument(Long val) {
		this.excelDocument = val;
	}

	@Input(required = Required.ALWAYS)
	@Name("LicenseKey")
	public void setLicenseKey(String val) {
		this.licenseKey = val;
	}

	@Input(required = Required.ALWAYS)
	@Name("documentName")
	public void setDocumentName(String val) {
		this.documentName = val;
	}

	@Input(required = Required.ALWAYS)
	@Name("saveInFolder")
	@FolderDataType
	public void setSaveInFolder(Long val) {
		this.saveInFolder = val;
	}

	@Name("newGeneratedDocument")
	public Long getNewGeneratedDocument() {
		return newGeneratedDocument;
	}

	@Name("errorOccured")
	public boolean getErrorOccured() {
		return errorOccured;
	}

	@Name("errorMessage")
	public String getErrorMessage() {
		return errorMessage;
	}

}
