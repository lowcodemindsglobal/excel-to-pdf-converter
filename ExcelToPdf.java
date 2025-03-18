package excelservice;

import java.io.FileInputStream;
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
import com.aspose.cells.License;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

@PaletteInfo(paletteCategory = "Appian Smart Services", palette = "Document Management")
public class ExcelToPdf extends AppianSmartService {

	private static final Logger LOG = Logger.getLogger(ExcelToPdf.class);
	@SuppressWarnings("unused")
	private final SmartServiceContext smartServiceCtx;
	private Long excelDocument;
	private Long licenseFile;
	private String documentName;
	private Long saveInFolder;
	private Long newGeneratedDocument;
	private final ContentService contentService;
	Long createdDocument;
	String filePath;
	private final String extensionValue = "pdf";
	String xlFileName;
	String licenseFileName;
	private boolean errorOccured;
	private String errorMessage;

	public void run() throws SmartServiceException {

		try {

			xlFileName = contentService.getInternalFilename(excelDocument);
			licenseFileName = contentService.getInternalFilename(licenseFile);

		} catch (InvalidContentException e1) {
			errorOccured = true;
			errorMessage = "InvalidContentException";
			LOG.error("InvalidContentException");
		}
		try {
			convertExcelToPdf(xlFileName, filePath, documentName, licenseFileName);

		} catch (Exception e) {
			LOG.error("Exception : " + e);
		}
	}

	public ExcelToPdf(SmartServiceContext smartServiceCtx, ContentService cs_) {
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

	public void convertExcelToPdf(String excelDocumentPath, String outputPath, String documentName, String licensePath)
			throws Exception {

		PdfSaveOptions pdfSaveOptions = null;
		Workbook workbook = null;
		Worksheet worksheet = null;

		FileInputStream fstream = new FileInputStream(licensePath);
		// Instantiate the License class
		License license = new License();
		// Set the license through the stream object
		license.setLicense(fstream);

		try {

			workbook = new Workbook(excelDocumentPath);

			for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
				worksheet = workbook.getWorksheets().get(i);

				worksheet.autoFitColumns();
			}

			pdfSaveOptions = new PdfSaveOptions();
			pdfSaveOptions.setOnePagePerSheet(true);

		} catch (Exception e) {
			errorOccured = true;
			errorMessage = "Invalid Sheet Name ERROR : ";

		} finally {

			if (fstream != null)
				try {
					fstream.close();
				} catch (Exception e) {
					errorOccured = true;
					errorMessage = "Invalid License File " + e;
					LOG.error("License File Error : " + e);
				}

		}

		if (errorOccured == false) {
			createdDocument = createDocument(documentName, extensionValue);
			filePath = contentService.getInternalFilename(createdDocument);
			workbook.save(filePath, pdfSaveOptions);
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
	@Name("licenseFile")
	@DocumentDataType
	public void setLicenseFile(Long val) {
		this.licenseFile = val;
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
