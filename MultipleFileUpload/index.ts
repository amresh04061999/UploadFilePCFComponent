import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class MultipleFileUpload
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _container: HTMLDivElement;
  private _notifyOutputChanged: () => void;
  private _fileSections: HTMLDivElement[] = [];
  private _fileCount: number = 0;
  private _maxFiles: number;
  private _chooseFileButton: HTMLButtonElement;
  private _dropZoneContainer: HTMLDivElement;
  private mainContainer: HTMLDivElement;
  private paragraph: HTMLDivElement;
  private _fileInput: HTMLInputElement;
  private _dropdown: HTMLSelectElement;
  private _files: {
    name: string;
    content: string;
    blobUrl: string;
    size: number;
  }[] = []; // Store files with blobUrl and size
  private _errorMessage: HTMLDivElement; // Error message element
  private _fileNamesContainer: HTMLDivElement; // Container for file names
  private _context: ComponentFramework.Context<IInputs>;
  private _currentValue: string[];
  constructor() {}
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._container = container;
    this._notifyOutputChanged = notifyOutputChanged;
    this._context = context;
    this._currentValue = this._currentValue || [];
    this._maxFiles =
      this._context.parameters.fileRangeValue?.raw !== null
        ? this._context.parameters.fileRangeValue.raw
        : 0;
    // Load external stylesheets and set up file upload UI
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href =
      "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css";
    document.head.appendChild(link);

  
    //    main div
    this.mainContainer = document.createElement("div");
    this.mainContainer.className = "upload-container";
    this._container.appendChild(this.mainContainer);

    const uploadHeader = document.createElement("div");
    uploadHeader.className = "upload-header";
    const title = document.createElement("h2");
    title.innerText = `Upload and attach files`;
    this.paragraph = document.createElement("p");
    const maximumFiles = document.createElement("p");
    maximumFiles.className = "maxFile";

    // const SaveButton = document.createElement("button");
    // SaveButton.innerText="saveFile"
    maximumFiles.innerText = `You can upload Maximum ${this._maxFiles} files.`;
    uploadHeader.appendChild(title);
    uploadHeader.appendChild(maximumFiles);
    uploadHeader.appendChild(this.paragraph);
    // uploadHeader.appendChild(SaveButton);
    this.mainContainer.appendChild(uploadHeader);
    this._errorMessage = document.createElement("div");
    this._errorMessage.className = "error-message";
    this.addFileSection();
  }
  private addFileSection(): void {
    const uploadBox = document.createElement("div");
    uploadBox.className = "upload-box";
    const fileSection = document.createElement("div");
    fileSection.className = "upload-drag";
    this._dropZoneContainer = document.createElement("div");
    const dropZone = document.createElement("div");
    this._dropZoneContainer.appendChild(dropZone);
    dropZone.className = "drop-zone";
    const containerDragDrop = document.createElement("div");
    containerDragDrop.className = "containerDragDrop";
    const uploadImage = document.createElement("div");
    uploadImage.className = "uploadImage";
    uploadImage.innerHTML = '<i class="fa fa-upload" aria-hidden="true"></i>';
    const setHeadingDropZone = document.createElement("div");
    const setHeadingDropZone1 = document.createElement("div");
    setHeadingDropZone.className = "headingDropZone";
    setHeadingDropZone.innerText = "Click to upload";
    setHeadingDropZone1.innerText = "Or Drag and Drop here";
    dropZone.appendChild(uploadImage);
    containerDragDrop.appendChild(setHeadingDropZone);
    containerDragDrop.appendChild(setHeadingDropZone1);
    dropZone.appendChild(containerDragDrop);
    setHeadingDropZone.addEventListener("click", () => {
      this._fileInput.click();
    });
    // Create the file input element for selecting files
    this._fileInput = document.createElement("input");
    this._fileInput.type = "file";
    // this._fileInput.accept=this._currentValue;
    this._fileInput.style.display = "none";
    console.log("This is current value", this._currentValue);

    // Create the "Choose Files" button
    this._chooseFileButton = document.createElement("button");
    this._chooseFileButton.className = "chooseFileButton";
    this._chooseFileButton.innerText = "Choose Files";
    this._chooseFileButton.title =
      "Select PDF files to attach. Upload only PDF files.";
    this._chooseFileButton.addEventListener("click", () => {
      this._fileInput.click();
    });

    // Create container for displaying selected files
    this._fileNamesContainer = document.createElement("div");
    this._fileNamesContainer.className = "file-names-container";

    // Drag and drop events
    dropZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropZone.classList.add("drop-zone-hover");
    });

    dropZone.addEventListener("dragleave", () => {
      dropZone.classList.remove("drop-zone-hover");
    });

    dropZone.addEventListener("drop", async (e) => {
      e.preventDefault();
      dropZone.classList.remove("drop-zone-hover");
      const droppedFiles = e.dataTransfer?.files;
      if (droppedFiles) {
        this.handleFiles(Array.from(droppedFiles), this._fileInput);
      }
    });

    // Handle file selection from the file input
    this._fileInput.addEventListener("change", async () => {
      const selectedFiles = this._fileInput.files;
      if (selectedFiles) {
        this.handleFiles(Array.from(selectedFiles), this._fileInput);
      }
    });

    // Append elements to the file section
    uploadBox.appendChild(fileSection);
    fileSection.appendChild(this._dropZoneContainer);
    fileSection.appendChild(this._chooseFileButton);
    this.mainContainer.appendChild(this._errorMessage);
    fileSection.appendChild(this._fileNamesContainer);
    fileSection.appendChild(this._fileInput);
    this.mainContainer.appendChild(fileSection);
    this._fileSections.push(fileSection);
  }

  // Handle selected or dropped files
  private async handleFiles(
    files: File[],
    fileInput: HTMLInputElement
  ): Promise<void> {
    let invalidFileSelected = false;
    let maxFilesExceeded = false;
    let filesToAdd = files;

    if (this._fileCount + filesToAdd.length > this._maxFiles) {
      maxFilesExceeded = true;
      filesToAdd = filesToAdd.slice(0, this._maxFiles - this._fileCount);
    }
    const fragment = document.createDocumentFragment();
    for (const file of filesToAdd) {
      if (this.isValidFile(file)) {
        const { base64Content, blobUrl, size } =
          await this.readFileAsBase64AndBlob(file);
        this._files.push({
          name: file.name,
          content: base64Content,
          blobUrl,
          size,
        });
        this._fileCount++;
        this.addFileDisplay(fragment, file, blobUrl, size);
      } else {
        invalidFileSelected = true;
      }
    }
    this._fileNamesContainer.appendChild(fragment);
    if (invalidFileSelected) {
      this.showError("Please select only PDF files.");
    } else if (maxFilesExceeded) {
      this.showError(`You can upload up to ${this._maxFiles} PDF files.`);
    } else {
      this.hideError();
    }
    fileInput.value = "";
    this.updateChooseFileButtonState();
    this._notifyOutputChanged();
  }
  private isValidFile(file: File): boolean {
    const validTypes = [
      "application/pdf",                          // .pdf
      "application/msword",                       // .doc
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // .docx
      "application/vnd.ms-powerpoint",           // .ppt
      "application/vnd.openxmlformats-officedocument.presentationml.presentation", // .pptx
      "application/vnd.ms-excel",                // .xls
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
      "image/jpeg",                               // .jpg, .jpeg
      "image/png",                                // .png
      "image/gif",                                // .gif
    ];
    return validTypes.includes(file.type);
  }
  private addFileDisplay(
    container: DocumentFragment,
    file: File,
    blobUrl: string,
    size: number
  ): void {
    const fileDisplay = document.createElement("div");
    fileDisplay.className = "file-display";
    const fileNameFileSizeContainer = document.createElement("div");
    fileNameFileSizeContainer.className = "fileNameFileSizeContainer";
    const fileName = document.createElement("a");
    const fileSize = document.createElement("div");
    fileSize.className = "fileSize";
    fileName.className = "file-name";
    fileName.innerText = `${this.truncateFileName(file.name)}`;
    fileSize.innerText = `${(size / 1024).toFixed(2)} KB`;
    fileNameFileSizeContainer.appendChild(fileName);
    fileNameFileSizeContainer.appendChild(fileSize);
    const detailsSection = document.createElement("div");
    detailsSection.className = "detailsSection";
    const ActionSection = document.createElement("div");
    const fileContainer = document.createElement("div");
    fileContainer.className = "fileContainer";
    const fileDiv = document.createElement("div");
    fileDiv.innerHTML = '<i class="fa fa-file" aria-hidden="true"></i>';
    fileDiv.className = "fileDiv";
    fileContainer.appendChild(fileDiv);
    const removeButton = document.createElement("button");
    const download = document.createElement("a");
    download.title = `${file.name} - ${size} bytes`;
    download.href = blobUrl;
    download.target = "_blank";
    download.download = file.name;
    download.style.marginRight = "10px";
    download.className = "download";
    download.innerHTML = '<i class="fas fa-download" aria-hidden="true"></i>';
    download.title = "Download this file";
    removeButton.className = "remove-button";
    removeButton.innerHTML = '<i class="fa fa-times" aria-hidden="true"></i>';
    removeButton.title = "Remove this file";
    const progressWrapper = document.createElement("div");
    progressWrapper.className = "progress-wrapper";
    // Create progress bar
    const progressBar = document.createElement("div");
    progressBar.className = "progress-bar";
    progressWrapper.appendChild(progressBar);

    // Create status text
    const statusText = document.createElement("p");
    statusText.className = "status-text";
    statusText.textContent = "Uploading... 0%";
    ActionSection.appendChild(download);
    ActionSection.appendChild(removeButton);
    removeButton.addEventListener("click", () => {
      fileDisplay.remove();
      this._files = this._files.filter((f) => f.name !== file.name);
      this._fileCount--;

      if (this._fileCount <= this._maxFiles) {
        this.hideError(); // Hide error message when within allowed limit
      }
      this.updateChooseFileButtonState();
      this._notifyOutputChanged();
    });

    detailsSection.appendChild(fileContainer);
    detailsSection.appendChild(fileNameFileSizeContainer);
    const listContainer = document.createElement("div");
    const progressWrapperContainer = document.createElement("div");
    const listItemContainer = document.createElement("div");
    progressWrapperContainer.className = "progressWrapperContainer";
    progressWrapperContainer.appendChild(progressWrapper);
    progressWrapperContainer.appendChild(statusText);
    listContainer.className = "listContainer";
    listContainer.appendChild(detailsSection);
    listContainer.appendChild(ActionSection);
    // listContainer.appendChild(progressWrapperContainer)
    fileDisplay.appendChild(listContainer);
    this._fileNamesContainer.appendChild(fileDisplay);
    this.mainContainer.appendChild(this._fileNamesContainer);
  }

  // Single read for both base64 and Blob URL
  private readFileAsBase64AndBlob(
    file: File
  ): Promise<{ base64Content: string; blobUrl: string; size: number }> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const base64Content = reader.result as string;
        const blob = new Blob([reader.result as ArrayBuffer], {
          type: file.type,
        });
        const blobUrl = URL.createObjectURL(blob);
        const size = file.size;
        resolve({ base64Content, blobUrl, size });
      };
      reader.onerror = (error) => reject(error);
      reader.readAsDataURL(file);
    });
  }

  private truncateFileName(fileName: string, maxLength: number = 30): string {
    return fileName.length > maxLength
      ? `${fileName.substr(0, maxLength - 3)}...`
      : fileName;
  }

  private updateChooseFileButtonState(): void {
    this._chooseFileButton.disabled = this._fileCount >= this._maxFiles;
  }

  private showError(message: string): void {
    this._errorMessage.innerText = message;
    this._errorMessage.style.display = "block";
  }

  private hideError(): void {
    this._errorMessage.innerText = "";
    this._errorMessage.style.display = "none";
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;
    const isFileActive = context.parameters.ChooseFileActive.raw;
    const DragAndDropActive = context.parameters.DragAndDropActive.raw;

    // Show or hide the button based on the value of ChooseFIleActive
    if (this._chooseFileButton) {
      if (isFileActive) {
        this._chooseFileButton.style.display = "block";
      } else {
        this._chooseFileButton.style.display = "none";
      }
    }
  
    // Show or hide the button based on the value of ChooseFIleActive
    if (this._dropZoneContainer) {
      if (DragAndDropActive) {
        this._dropZoneContainer.style.display = "block";
      } else {
        this._dropZoneContainer.style.display = "none";
      }
    }
    const allowMultipleFiles = context.parameters.AllowMultipleFiles.raw;
    if (this._fileInput) {
      if (allowMultipleFiles) {
        this._fileInput.multiple = true;
      } else {
        this._fileInput.multiple = false;
      }
    }
    const updateFileSupport = (isActive: boolean, extensions: string[]) => {
      extensions.forEach((ext) => {
        if (isActive) {
          if (!this._currentValue.includes(ext)) {
            this._currentValue.push(ext); // Add unique extensions
          }
        } else {
          this._currentValue = this._currentValue.filter((item) => item !== ext); // Remove extensions if inactive
        }
      });
    
      const acceptString = this._currentValue.join(",");
      this.paragraph.innerText = `Supported Files: ${acceptString}`;
      this._fileInput.accept = acceptString;
    };
    
    // Dynamically update supported files
    updateFileSupport(context.parameters.pdfFileFormateActive?.raw, [".pdf"]);
    updateFileSupport(context.parameters.imageFileFormateActive?.raw, [
      "image/gif",
      "image/jpeg",
      "image/png",
    ]);
    updateFileSupport(context.parameters.wordFileFormateActive?.raw, [".doc", ".docx"]);
    updateFileSupport(context.parameters.pptFileFormateActive?.raw, [".ppt", ".pptx"]);
    updateFileSupport(context.parameters.excelFileFormateActive?.raw, [".xls", ".xlsx"]);
    

   
  }
  public getOutputs(): IOutputs {
    return {
      // dropdownValue: this._currentValue,
      fileCount: this._fileCount,
      files: JSON.stringify(this._files),
    };
  }

  public destroy(): void {
    this._dropdown.removeEventListener("change", () => {});
    this._fileSections.forEach((section) =>
      this._container.removeChild(section)
    );
    this._fileSections = [];
  }
}
