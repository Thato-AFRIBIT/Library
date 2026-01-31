import { Version } from '@microsoft/sp-core-library';
import { 
  BaseClientSideWebPart, 
  IPropertyPaneConfiguration, 
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneButton,
  IPropertyPaneGroup,
  IPropertyPaneField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import styles from './LibraryWebPart.module.scss';

export interface IResourceSection {
  id: string;
  sectionTitle: string;
  libraryName: string;
  folderName: string;
}

export interface IUnifiedResourcesWebPartProps {
  pageTitle: string;
  mainLibraryName: string;
  mainFolderName: string;
  sections: IResourceSection[];
  helpArticleHtml: string;
  helpImageUrl: string;
}

interface IFileItem {
  FileLeafRef: string;
  FileRef: string;
  ServerRelativeUrl: string;
  FileType: string;
  FolderPath?: string;
}

interface IListInfo {
  Title: string;
  Id: string;
  BaseTemplate: number;
}

interface IFolderInfo {
  Name: string;
  ServerRelativeUrl: string;
}

interface IFolderItemResponse {
  Id: number;
  Title?: string;
  FileLeafRef: string;
  FileRef: string;
  FSObjType: number;
}

interface IFileItemResponse {
  Id: number;
  Title: string;
  File: {
    Name: string;
    ServerRelativeUrl: string;
    TimeCreated: string;
    Length: number;
  };
}

// Help Article Dialog
class ArticleDialog {
  public content: string = "";
  public image: string = "";
  private _dialogElement: HTMLElement | null = null;
  private _overlayElement: HTMLElement | null = null;

  private _formatText(text: string): string {
    if (!text) return "";
    let formatted = text;
    formatted = formatted.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    formatted = formatted.replace(/^\s*[*|-]\s+(.*)$/gm, '<div style="margin-left:20px; display: list-item; list-style-type: disc;">$1</div>');
    formatted = formatted.replace(/^\s*(\d+)\.\s+(.*)$/gm, '<div style="margin-left:20px; display: list-item; list-style-type: decimal;">$2</div>');
    return formatted.replace(/\n/g, '<br/>');
  }

  private render(): void {
    // Create overlay
    this._overlayElement = document.createElement('div');
    this._overlayElement.style.position = 'fixed';
    this._overlayElement.style.top = '0';
    this._overlayElement.style.left = '0';
    this._overlayElement.style.width = '100%';
    this._overlayElement.style.height = '100%';
    this._overlayElement.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    this._overlayElement.style.zIndex = '1000';
    this._overlayElement.addEventListener('click', () => this.close());

    // Create dialog
    this._dialogElement = document.createElement('div');
    this._dialogElement.innerHTML = `
      <div class="${styles.articleContainer}">
        <div class="${styles.articleHeader}">
           <h3>Help Guide</h3>
           <button id="closeDialogBtn" style="cursor: pointer;">Close</button>
        </div>
        <div class="${styles.articleBody}">
          ${this.image ? `<img src="${this.image}" class="${styles.articleImage}"/>` : ''}
          <div class="${styles.articleContent}">
            ${this._formatText(this.content)}
          </div>
        </div>
      </div>`;
    
    this._dialogElement.style.position = 'fixed';
    this._dialogElement.style.top = '50%';
    this._dialogElement.style.left = '50%';
    this._dialogElement.style.transform = 'translate(-50%, -50%)';
    this._dialogElement.style.zIndex = '1001';
    this._dialogElement.style.maxHeight = '90vh';
    this._dialogElement.style.overflowY = 'auto';

    const closeBtn = this._dialogElement.querySelector('#closeDialogBtn');
    if (closeBtn) {
      closeBtn.addEventListener('click', (e: Event) => { 
        e.stopPropagation();
        this.close(); 
      });
    }
  }

  public show(): Promise<void> {
    return new Promise((resolve) => {
      this.render();
      
      if (this._overlayElement) {
        document.body.appendChild(this._overlayElement);
      }
      if (this._dialogElement) {
        document.body.appendChild(this._dialogElement);
      }
      
      resolve();
    });
  }

  public close(): void {
    if (this._overlayElement && this._overlayElement.parentNode) {
      document.body.removeChild(this._overlayElement);
    }
    if (this._dialogElement && this._dialogElement.parentNode) {
      document.body.removeChild(this._dialogElement);
    }
  }
}

export default class UnifiedResourcesWebPart extends BaseClientSideWebPart<IUnifiedResourcesWebPartProps> {

  private _propertyPaneObserver: MutationObserver | null = null;
  private _addLinkModal: HTMLDivElement | null = null;

  // Section-specific data storage (including main section which is 'main')
  private sectionData: { [sectionId: string]: {
    allFiles: IFileItem[];
    displayedFiles: IFileItem[];
    isLoading: boolean;
    hasMoreFiles: boolean;
    showBackButton: boolean;
    currentFolder: string;
    availableFolders: IFolderInfo[];
  }} = {};

  private getBatchSize(): number {
    const totalSections = 1 + (this.properties.sections ? this.properties.sections.length : 0);
    if (totalSections === 1) return 20;
    if (totalSections === 2) return 15;
    if (totalSections === 3) return 10;
    if (totalSections >= 4) return 5;
    return 20;
  }

  private libraryOptions: IPropertyPaneDropdownOption[] = [];
  private folderOptions: { [libraryName: string]: IPropertyPaneDropdownOption[] } = {};
  private isLibraryOptionsLoading: boolean = false;
  private isFolderOptionsLoading: { [libraryName: string]: boolean } = {};

  private readonly allowedFileTypes: string[] = [
    'pdf', 'doc', 'docx', 'txt', 'xls', 'xlsx', 'ppt', 'pptx', 'odt', 'url'
  ];

  private readonly DOCUMENT_LIBRARY_TEMPLATE = 101;

  private _initializeSectionData(sectionId: string): void {
    if (!this.sectionData[sectionId]) {
      this.sectionData[sectionId] = {
        allFiles: [],
        displayedFiles: [],
        isLoading: false,
        hasMoreFiles: true,
        showBackButton: true,
        currentFolder: '',
        availableFolders: []
      };
    }
  }

  public render(): void {
    // Initialize properties
    if (!this.properties.sections || !Array.isArray(this.properties.sections)) {
      this.properties.sections = [];
    }
    if (!this.properties.pageTitle) {
      this.properties.pageTitle = 'Resources';
    }

    // Check if main section is configured
    const hasMainSection = this.properties.mainLibraryName && this.properties.mainLibraryName.trim() !== '';

    // Calculate responsive max height
    const totalSections = (hasMainSection ? 1 : 0) + this.properties.sections.length;
    let maxHeight = '650px';
    if (totalSections === 2) maxHeight = '400px';
    else if (totalSections === 3) maxHeight = '300px';
    else if (totalSections >= 4) maxHeight = '250px';

    // Build HTML
    this.domElement.innerHTML = `
    <div class="${styles.container}">
      <div class="${styles.header}">
        <!-- Back button on the left -->
        <button class="${styles.backButton}" id="backButton" title="Go to Home">
          <svg version="1.1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 512 512" class="${styles.backIcon}">
            <polygon points="512,228.19 106.42,228.19 253.6,81.02 214.3,41.7 0,256 214.3,470.3 253.6,430.98 106.42,283.81 512,283.81"/>
          </svg>
          <span class="${styles.backText}">Back</span>
        </button>
        
        <!-- Page Title centered in the middle -->
        <div class="${styles.headerTitle}">${this._escapeHtml(this.properties.pageTitle)}</div>
        
        <!-- Right-side buttons -->
        <div class="${styles.headerActions}">
          <button class="${styles.helpButton}" id="helpButton" title="Help Guide">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <circle cx="12" cy="12" r="10"></circle>
              <path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"></path>
              <line x1="12" y1="17" x2="12.01" y2="17"></line>
            </svg>
            <span>Help</span>
          </button>
        </div>
      </div>
      
      ${hasMainSection ? `
      <!-- Main Section (Page Header Library) -->
      <div class="${styles.section}" data-section-id="main">
        <div class="folderList-main" data-section-id="main"></div>
        <div id="fileList-main" class="${styles.fileList}" style="max-height: ${maxHeight};"></div>
        <div id="loading-main" class="${styles.loading}" style="display:none;">Loading files...</div>
      </div>
      ` : ''}
      
      ${this.properties.sections.map(section => `
      <!-- Additional Section -->
      <div class="${styles.section}" data-section-id="${section.id}">
        <div class="${styles.sectionHeader}">
          <h2 class="${styles.sectionTitle}">${this._escapeHtml(section.sectionTitle || 'Additional Resources')}</h2>
        </div>
        <div class="folderList-${section.id}" data-section-id="${section.id}"></div>
        <div id="fileList-${section.id}" class="${styles.fileList}" style="max-height: ${maxHeight};"></div>
        <div id="loading-${section.id}" class="${styles.loading}" style="display:none;">Loading files...</div>
      </div>
      `).join('')}
      
      ${!hasMainSection && this.properties.sections.length === 0 ? `
      <div class="${styles.empty}">
        <strong>Setup Required</strong><br><br>
        To display documents, please configure the web part settings:<br><br>
        1. Click the <strong>Edit</strong> button in the top right corner<br>
        2. Select a <strong>Document Library</strong> for the main page section<br>
        3. Optionally add more sections by clicking <strong>+ Add Section</strong><br><br>
        <em>You can configure up to 3 additional sections</em>
      </div>
      ` : ''}
    </div>`;

    // Add event listeners
    const backButton = this.domElement.querySelector("#backButton") as HTMLButtonElement;
    if (backButton) {
      backButton.addEventListener("click", () => this._handleBackButton());
    }

    const helpButton = this.domElement.querySelector("#helpButton") as HTMLButtonElement;
    if (helpButton) {
      helpButton.addEventListener("click", () => this._showHelpDialog());
    }

    const addLinkButton = this.domElement.querySelector("#addLinkButton") as HTMLButtonElement;
    if (addLinkButton) {
      addLinkButton.addEventListener("click", () => this._showAddLinkModal());
    }

    // Load main section if configured
    if (hasMainSection) {
      this._initializeSectionData('main');
      this._attachFolderEventListeners('main');
      this._loadFilesForSection('main', this.properties.mainLibraryName, this.properties.mainFolderName)
        .catch((err) => console.error(`Error loading main section:`, err));
    }

    // Load additional sections
    this.properties.sections.forEach(section => {
      this._initializeSectionData(section.id);
      this._attachFolderEventListeners(section.id);
      this._loadFilesForSection(section.id, section.libraryName, section.folderName)
        .catch((err) => console.error(`Error loading section ${section.id}:`, err));
    });
  }

  private _showHelpDialog(): void {
    const dialog = new ArticleDialog();
    dialog.content = this.properties.helpArticleHtml || "No help content configured. Please add content in the property pane.";
    dialog.image = this.properties.helpImageUrl;
    dialog.show().catch((e: Error) => console.error('Error showing help dialog:', e));
  }

  private _showAddLinkModal(): void {
    if (this._addLinkModal) {
      document.body.removeChild(this._addLinkModal);
    }

    this._addLinkModal = document.createElement('div');
    this._addLinkModal.className = styles.modalOverlay;
    
    const libraryOptionsHtml = this.libraryOptions.map(lib => `
      <option value="${this._escapeHtml(String(lib.key))}">${this._escapeHtml(String(lib.text))}</option>
    `).join('') || '<option value="">No libraries available</option>';

    this._addLinkModal.innerHTML = `
      <div class="${styles.modal}">
        <div class="${styles.modalHeader}">
          <h3 style="font-family: Arial, sans-serif;">Add Link</h3>
          <button class="${styles.modalClose}" id="closeModal">&times;</button>
        </div>
        <div class="${styles.modalBody}">
          <div class="${styles.formGroup}">
            <label for="linkUrl" style="font-family: Arial, sans-serif;">Link URL *</label>
            <input type="url" id="linkUrl" class="${styles.formInput}" 
                   placeholder="https://example.com" required>
            <div class="${styles.formHint}">Enter the full URL including https://</div>
          </div>
          
          <div class="${styles.formGroup}">
            <label for="linkTitle" style="font-family: Arial, sans-serif;">Link Title *</label>
            <input type="text" id="linkTitle" class="${styles.formInput}" 
                   placeholder="Link Title" required>
            <div class="${styles.formHint}">A descriptive name for this link</div>
          </div>
          
          <div class="${styles.formGroup}">
            <label for="targetLibrary" style="font-family: Arial, sans-serif;">Save to Library *</label>
            <select id="targetLibrary" class="${styles.formSelect}" required>
              <option value="">Select a library...</option>
              ${libraryOptionsHtml}
            </select>
          </div>
          
          <div class="${styles.formGroup}">
            <label for="targetFolder" style="font-family: Arial, sans-serif;">Folder (Optional)</label>
            <select id="targetFolder" class="${styles.formSelect}">
              <option value="">Root folder (default)</option>
            </select>
            <div class="${styles.formHint}">Select a folder within the library (optional)</div>
          </div>
          
          <div class="${styles.modalActions}">
            <button type="button" class="${styles.cancelButton}" id="cancelButton">Cancel</button>
            <button type="button" class="${styles.saveButton}" id="saveLinkButton">Save Link</button>
          </div>
          
          <div id="linkStatus" class="${styles.statusMessage}" style="display: none;"></div>
        </div>
      </div>
    `;

    document.body.appendChild(this._addLinkModal);

    const closeButton = this._addLinkModal.querySelector('#closeModal') as HTMLButtonElement;
    const cancelButton = this._addLinkModal.querySelector('#cancelButton') as HTMLButtonElement;
    const saveButton = this._addLinkModal.querySelector('#saveLinkButton') as HTMLButtonElement;
    const linkUrlInput = this._addLinkModal.querySelector('#linkUrl') as HTMLInputElement;
    const linkTitleInput = this._addLinkModal.querySelector('#linkTitle') as HTMLInputElement;
    const targetLibrarySelect = this._addLinkModal.querySelector('#targetLibrary') as HTMLSelectElement;
    const targetFolderSelect = this._addLinkModal.querySelector('#targetFolder') as HTMLSelectElement;

    targetLibrarySelect.addEventListener('change', async () => {
      const libraryName = targetLibrarySelect.value;
      if (libraryName) {
        await this._loadFolderOptions(libraryName);
        const folders = this.folderOptions[libraryName] || [{ key: '', text: 'Root folder (default)' }];
        targetFolderSelect.innerHTML = folders.map(folder => 
          `<option value="${String(folder.key)}">${folder.text}</option>`
        ).join('');
      } else {
        targetFolderSelect.innerHTML = '<option value="">Root folder (default)</option>';
      }
    });

    const closeModal = (): void => {
      if (this._addLinkModal) {
        document.body.removeChild(this._addLinkModal);
        this._addLinkModal = null;
      }
    };

    closeButton.addEventListener('click', closeModal);
    cancelButton.addEventListener('click', closeModal);

    this._addLinkModal.addEventListener('click', (e) => {
      if (e.target === this._addLinkModal) closeModal();
    });

    saveButton.addEventListener('click', async () => {
      const url = linkUrlInput.value.trim();
      const title = linkTitleInput.value.trim();
      const libraryName = targetLibrarySelect.value;
      const folderName = targetFolderSelect.value;

      if (!url) {
        this._showStatus('Please enter a URL', 'error');
        return;
      }
      if (!title) {
        this._showStatus('Please enter a title', 'error');
        return;
      }
      if (!libraryName) {
        this._showStatus('Please select a library', 'error');
        return;
      }

      try {
        // eslint-disable-next-line no-new
        new URL(url);
      } catch {
        this._showStatus('Please enter a valid URL (include https://)', 'error');
        return;
      }

      const library = this.libraryOptions.find((lib: IPropertyPaneDropdownOption) => String(lib.key) === libraryName);
      if (!library) {
        this._showStatus('Selected library not found', 'error');
        return;
      }

      saveButton.disabled = true;
      saveButton.textContent = 'Saving...';
      this._showStatus('Saving link...', 'info');

      try {
        await this._createLinkFile(url, title, libraryName, folderName);
        this._showStatus('✓ Link saved successfully!', 'success');
        
        // Refresh main section if applicable
        if (this.properties.mainLibraryName === libraryName) {
          this._loadFilesForSection('main', this.properties.mainLibraryName, this.properties.mainFolderName).catch(console.error);
        }
        
        // Refresh matching additional sections
        this.properties.sections.forEach(section => {
          if (section.libraryName === libraryName) {
            this._loadFilesForSection(section.id, section.libraryName, section.folderName).catch(console.error);
          }
        });
        
        setTimeout(() => closeModal(), 2000);
        
      } catch (error) {
        console.error('Error saving link:', error);
        this._showStatus(`✗ Error: ${(error as Error).message}`, 'error');
        saveButton.disabled = false;
        saveButton.textContent = 'Save Link';
      }
    });

    setTimeout(() => linkUrlInput.focus(), 100);
  }

  private _showStatus(message: string, type: 'info' | 'success' | 'error'): void {
    if (!this._addLinkModal) return;
    
    const statusElement = this._addLinkModal.querySelector('#linkStatus') as HTMLDivElement;
    if (statusElement) {
      statusElement.textContent = message;
      statusElement.className = styles.statusMessage;
      const statusClass = `status${type.charAt(0).toUpperCase() + type.slice(1)}`;
      const className = (styles as { [key: string]: string })[statusClass];
      if (className) {
        statusElement.classList.add(className);
      }
      statusElement.style.display = 'block';
      
      if (type === 'success') {
        setTimeout(() => {
          if (statusElement.parentNode) statusElement.style.display = 'none';
        }, 5000);
      }
    }
  }

  private async _createLinkFile(url: string, title: string, libraryName: string, folderName?: string): Promise<void> {
    const cleanFileName = title
      .replace(/[<>:"\\|?*]/g, '-')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, 100);

    if (!cleanFileName) throw new Error('Invalid title provided');

    const fileContent = `[InternetShortcut]\r\nURL=${url}`;
    
    let folderPath = `${this.context.pageContext.web.serverRelativeUrl}/${libraryName}`;
    if (folderName && folderName.trim() !== '') {
      folderPath = `${folderPath}/${folderName}`;
    }
    folderPath = folderPath.replace(/\/\//g, '/');
    
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const digest = await this._getRequestDigest();
    
    const apiUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files/Add(url='${encodeURIComponent(cleanFileName + '.url')}',overwrite=true)`;
    
    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
        'Content-Type': 'application/octet-stream'
      },
      body: fileContent
    });
    
    if (!response.ok) {
      let errorMessage = `Failed to save link: ${response.status} ${response.statusText}`;
      if (response.status === 403) {
        errorMessage += '\n\nYou may not have permission to add files to this library, or .url files may be blocked.';
      } else if (response.status === 404) {
        errorMessage += '\n\nThe library or folder path may not exist.';
      }
      throw new Error(errorMessage);
    }
  }

  private async _getRequestDigest(): Promise<string> {
    try {
      if (this.context.pageContext.legacyPageContext) {
        const digest = (this.context.pageContext.legacyPageContext as Record<string, unknown>).formDigestValue;
        if (digest) return digest as string;
      }

      const digestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`;
      const response = await this.context.spHttpClient.post(
        digestUrl,
        SPHttpClient.configurations.v1,
        {}
      );

      if (response.ok) {
        const data = await response.json();
        return data.FormDigestValue;
      }

      throw new Error('Failed to get request digest');
    } catch (error) {
      console.error('Error getting request digest:', error);
      throw error;
    }
  }

  private _renderFolderList(sectionId: string): string {
    const section = this.sectionData[sectionId];
    if (section.showBackButton || section.availableFolders.length === 0) {
      return '';
    }

    return `
      <div class="${styles.folderList}" id="folderList-${sectionId}" data-section-id="${sectionId}">
        <h3 class="${styles.folderTitle}">Folders</h3>
        ${section.availableFolders.map(folder => `
          <div class="${styles.folderItem}" data-folder="${this._escapeHtml(folder.Name)}" data-section-id="${sectionId}">
            <div class="${styles.folderIcon}">
              <svg xmlns="http://www.w3.org/2000/svg" height="20" viewBox="0 -960 960 960" width="20">
                <path d="M160-160q-33 0-56.5-23.5T80-240v-480q0-33 23.5-56.5T160-800h240l80 80h320q33 0 56.5 23.5T880-640v400q0-33-23.5-56.5T800-160H160Z"/>
              </svg>
            </div>
            <div class="${styles.folderName}">${this._escapeHtml(folder.Name)}</div>
            <div class="${styles.folderArrow}">
              <svg xmlns="http://www.w3.org/2000/svg" height="20" viewBox="0 -960 960 960" width="20">
                <path d="m504-480-156 156q-11 11-11 28t11 28q11 11 28 11t28-11l184-184q6-6 8.5-13t2.5-15q0-8-2.5-15t-8.5-13L404-692q-11-11-28-11t-28 11q-11 11-11 28t11 28l156 156Z"/>
              </svg>
            </div>
          </div>
        `).join('')}
      </div>
    `;
  }

  private _renderFoldersIntoDOM(sectionId: string): void {
    const folderContainer = this.domElement.querySelector(`.folderList-${sectionId}`) as HTMLElement;
    if (!folderContainer) return;
    
    folderContainer.innerHTML = this._renderFolderList(sectionId);
    this._attachFolderEventListeners(sectionId);
  }

  private _attachFolderEventListeners(sectionId: string): void {
    const section = this.sectionData[sectionId];
    if (!section || section.showBackButton || section.availableFolders.length === 0) {
      return;
    }

    const folderItems = this.domElement.querySelectorAll(`.${styles.folderItem}[data-section-id="${sectionId}"]`);
    folderItems.forEach(folderItem => {
      folderItem.addEventListener("click", () => {
        const folderName = folderItem.getAttribute('data-folder');
        if (folderName) {
          this._handleFolderClick(sectionId, folderName);
        }
      });
    });
  }

  private _handleBackButton(): void {
    window.location.href = 'https://kzndard.sharepoint.com/sites/KZNDARDCentral';
  }

  private _handleFolderClick(sectionId: string, folderName: string): void {
    const section = this.sectionData[sectionId];
    section.showBackButton = true;
    section.currentFolder = folderName;
    this.render();
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    await this._loadLibraryOptions();
    this._setupPropertyPaneButtonStyling();
  }

  private _setupPropertyPaneButtonStyling(): void {
    try {
      this._propertyPaneObserver = new MutationObserver(() => {
        const pane = document.querySelector('.ms-PropertyPane');
        if (!pane) return;

        const buttons = pane.querySelectorAll('button.ms-Button, button[class*="ms-Button"]');
        buttons.forEach((btn) => {
          const htmlBtn = btn as HTMLElement;
          const text = (btn.textContent || '').trim();
          if (!text) return;

          if ((text as string).includes('Add Section')) {
            htmlBtn.style.backgroundColor = '#cfae70';
            htmlBtn.style.borderColor = '#cfae70';
            htmlBtn.style.color = '#000000';
            htmlBtn.style.height = '44px';
            htmlBtn.style.minHeight = '44px';
            htmlBtn.style.borderRadius = '6px';
            htmlBtn.style.padding = '12px 20px';
            htmlBtn.style.fontWeight = '600';
          }

          if ((text as string).startsWith('Remove Section')) {
            htmlBtn.style.backgroundColor = '#000000';
            htmlBtn.style.borderColor = '#000000';
            htmlBtn.style.color = '#ffffff';
            htmlBtn.style.height = '44px';
            htmlBtn.style.minHeight = '44px';
            htmlBtn.style.borderRadius = '6px';
            htmlBtn.style.padding = '12px 20px';
            htmlBtn.style.fontWeight = '600';
            
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const oldClick = (htmlBtn as any)._originalClick;
            if (!oldClick) {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              (htmlBtn as any)._originalClick = htmlBtn.onclick;
              htmlBtn.onclick = (e: MouseEvent) => {
                e.preventDefault();
                e.stopPropagation();
                
                const confirmDelete = confirm('Delete this section?');
                if (confirmDelete) {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  if ((htmlBtn as any)._originalClick) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    (htmlBtn as any)._originalClick.call(htmlBtn, e);
                  }
                }
                return false;
              };
            }
          }
        });
      });

      this._propertyPaneObserver.observe(document.body, { childList: true, subtree: true });
    } catch (error) {
      console.warn('Property pane button styling observer failed', error);
    }
  }

  protected onDispose(): void {
    if (this._propertyPaneObserver) {
      this._propertyPaneObserver.disconnect();
      this._propertyPaneObserver = null;
    }
    
    if (this._addLinkModal) {
      document.body.removeChild(this._addLinkModal);
    }
    
    super.onDispose();
  }

  private async _loadLibraryOptions(): Promise<void> {
    this.isLibraryOptionsLoading = true;
    
    try {
      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title,Id,BaseTemplate&$filter=BaseTemplate eq ${this.DOCUMENT_LIBRARY_TEMPLATE} and Hidden eq false`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        this.libraryOptions = data.value.map((list: IListInfo) => ({
          key: list.Title,
          text: list.Title
        }));
      }
    } catch (err) {
      console.error('Error loading libraries:', err);
      this.libraryOptions = [];
    } finally {
      this.isLibraryOptionsLoading = false;
    }
  }

  private async _loadFolderOptions(libraryName: string): Promise<void> {
    if (!libraryName) {
      this.folderOptions[libraryName] = [];
      return;
    }

    if (this.isFolderOptionsLoading[libraryName]) return;
    if (this.folderOptions[libraryName]) return;

    this.isFolderOptionsLoading[libraryName] = true;
    
    try {
      const folders = await this._getFoldersDirectQuery(libraryName);
      
      this.folderOptions[libraryName] = [
        { key: '', text: 'Root folder (default)' },
        ...folders.map((folder: IFolderInfo) => ({
          key: folder.Name,
          text: folder.Name
        }))
      ];

    } catch (err) {
      console.error('Error loading folders:', err);
      this.folderOptions[libraryName] = [{ key: '', text: 'Root folder (default)' }];
    } finally {
      this.isFolderOptionsLoading[libraryName] = false;
      this.context.propertyPane.refresh();
    }
  }

  private async _getFoldersDirectQuery(libraryName: string): Promise<IFolderInfo[]> {
    try {
      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${libraryName}')/items?$select=Id,Title,FileLeafRef,FileRef,FSObjType&$filter=FSObjType eq 1&$orderby=FileLeafRef&$top=1000`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Folder query failed: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();

      if (!data.value || data.value.length === 0) {
        return [];
      }

      const folders: IFolderInfo[] = data.value.map((folder: IFolderItemResponse) => ({
        Name: folder.FileLeafRef || folder.Title || '',
        ServerRelativeUrl: folder.FileRef
      }));

      return folders;

    } catch (error) {
      console.error('Error in direct folder query:', error);
      
      try {
        return await this._getFoldersEnumeration(libraryName);
      } catch (fallbackError) {
        console.error('Fallback also failed:', fallbackError);
        return [];
      }
    }
  }

  private async _getFoldersEnumeration(libraryName: string): Promise<IFolderInfo[]> {
    try {
      const libraryServerRelativeUrl = `${this.context.pageContext.web.serverRelativeUrl}/${libraryName}`;
      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(libraryServerRelativeUrl)}')/folders?$select=Name,ServerRelativeUrl&$top=1000`;
      
      const response = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      
      if (!response.ok) {
        throw new Error(`Folder enumeration failed: ${response.status}`);
      }

      const data = await response.json();
      return data.value || [];

    } catch (error) {
      console.error('Folder enumeration error:', error);
      throw error;
    }
  }

  private async _loadFilesForSection(sectionId: string, libraryName: string, folderName?: string): Promise<void> {
    const section = this.sectionData[sectionId];
    
    if (!libraryName || section.isLoading) return;

    section.isLoading = true;

    const loadingElement = this.domElement.querySelector(`#loading-${sectionId}`) as HTMLElement;
    if (loadingElement) {
      loadingElement.style.display = "block";
    }

    try {
      let allFiles = await this._getAllFilesFromLibrary(libraryName);

      const activeFolder = section.currentFolder || folderName;
      
      if (activeFolder && activeFolder.trim() !== '') {
        allFiles = this._filterFilesByFolder(allFiles, activeFolder);
        section.showBackButton = true;
        section.currentFolder = activeFolder;
      } else {
        section.showBackButton = true;
        section.currentFolder = '';
      }

      section.allFiles = allFiles.filter((file: IFileItem) => {
        return this.allowedFileTypes.indexOf(file.FileType.toLowerCase()) > -1;
      });

      const batchSize = this.getBatchSize();
      section.displayedFiles = section.allFiles.slice(0, batchSize);
      section.hasMoreFiles = section.allFiles.length > section.displayedFiles.length;

      this._renderFilesForSection(sectionId);

    } catch (err) {
      console.error(`Error loading documents for section ${sectionId}:`, err);
      this._renderErrorForSection(sectionId, `Error loading documents: ${(err as Error).message}`);
    } finally {
      section.isLoading = false;
      const loadingElement = this.domElement.querySelector(`#loading-${sectionId}`) as HTMLElement;
      if (loadingElement) {
        loadingElement.style.display = "none";
      }
    }
  }

  private async _getAllFilesFromLibrary(libraryName: string): Promise<IFileItem[]> {
    const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${libraryName}')/items?$select=Id,Title,File/Name,File/ServerRelativeUrl,File/TimeCreated,File/Length&$expand=File&$filter=FSObjType eq 0&$top=1000`;
    
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      apiUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}. Please check if the library name "${libraryName}" is correct.`);
    }

    const data = await response.json();
    
    return data.value
      .filter((item: IFileItemResponse) => item.File)
      .map((item: IFileItemResponse) => ({
        FileLeafRef: item.File.Name,
        FileRef: item.File.ServerRelativeUrl,
        ServerRelativeUrl: item.File.ServerRelativeUrl,
        FileType: this._getFileExtension(item.File.Name),
        FolderPath: this._extractFolderPath(item.File.ServerRelativeUrl, libraryName)
      }));
  }

  private _extractFolderPath(serverRelativeUrl: string, libraryName: string): string {
    const libraryPath = `${this.context.pageContext.web.serverRelativeUrl}/${libraryName}`.toLowerCase();
    const filePath = serverRelativeUrl.toLowerCase();
    
    const relativePath = filePath.replace(libraryPath, '').replace(/^\//, '');
    const pathParts = relativePath.split('/');
    
    return pathParts.slice(0, -1).join('/');
  }

  private _filterFilesByFolder(files: IFileItem[], folderName: string): IFileItem[] {
    const lowerFolderName = folderName.toLowerCase().trim();
    
    return files.filter(file => {
      const folderPath = file.FolderPath || '';
      const pathParts = folderPath.split('/');
      return (pathParts as string[]).includes(lowerFolderName);
    });
  }

  private _renderErrorForSection(sectionId: string, message: string): void {
    const fileListEl = this.domElement.querySelector(`#fileList-${sectionId}`) as HTMLElement;
    if (fileListEl) {
      fileListEl.innerHTML = `
        <div class="${styles.error}">
          <strong>Unable to load documents</strong><br><br>
          ${message}<br><br>
          Please try the following:<br>
          • Verify the library name is correct<br>
          • Check your permissions to access this library<br>
          • Refresh the page and try again<br>
          • <strong>Contact your administrator</strong> if the issue continues
        </div>
      `;
    }
  }

  private _loadMoreFilesForSection(sectionId: string): void {
    const section = this.sectionData[sectionId];
    if (section.isLoading || !section.hasMoreFiles) return;
    section.isLoading = true;

    const loadingElement = this.domElement.querySelector(`#loading-${sectionId}`) as HTMLElement;
    if (loadingElement) {
      loadingElement.style.display = "block";
    }

    setTimeout(() => {
      try {
        const batchSize = this.getBatchSize();
        const nextBatchStart = section.displayedFiles.length;
        const nextBatchEnd = nextBatchStart + batchSize;
        const nextBatch = section.allFiles.slice(nextBatchStart, nextBatchEnd);
        
        section.displayedFiles = section.displayedFiles.concat(nextBatch);
        section.hasMoreFiles = section.displayedFiles.length < section.allFiles.length;

        this._renderFilesForSection(sectionId);

      } catch (err) {
        console.error(`Error loading more documents for section ${sectionId}:`, err);
      } finally {
        section.isLoading = false;
        const loadingElement = this.domElement.querySelector(`#loading-${sectionId}`) as HTMLElement;
        if (loadingElement) {
          loadingElement.style.display = "none";
        }
      }
    }, 500);
  }

  private _renderFilesForSection(sectionId: string): void {
    const fileListEl = this.domElement.querySelector(`#fileList-${sectionId}`) as HTMLElement;
    if (!fileListEl) return;
    
    const section = this.sectionData[sectionId];
    fileListEl.innerHTML = '';

    this._renderFoldersIntoDOM(sectionId);

    if (section.displayedFiles.length === 0) {
      fileListEl.innerHTML = `
        <div class="${styles.empty}">
          <strong>No documents available</strong><br><br>
          There are currently no documents in this section.<br><br>
          <em>Supported formats: PDF, Word, Excel, PowerPoint, Text files, Web links</em>
        </div>
      `;
      return;
    }

    section.displayedFiles.forEach((item: IFileItem) => {
      const fileItem = document.createElement("div");
      fileItem.className = styles.fileItem;

      const infoHtml = `
        <div class="${styles.fileInfo}">
          <span class="${styles.fileType}">${this._getFileTypeDisplay(item.FileType)}</span>
          <span class="${styles.name}">${this._escapeHtml(item.FileLeafRef)}</span>
        </div>
      `;

      fileItem.innerHTML = infoHtml;

      const actions = document.createElement('div');
      actions.className = styles.actions;

      const isUrlShortcut = item.FileType.toLowerCase() === 'url';

      const openBtn = document.createElement('a');
      openBtn.className = `${styles.actionButton} ${styles.openButton}`;
      openBtn.title = isUrlShortcut ? 'Open target link' : 'Open in new tab';
      openBtn.setAttribute('rel', 'noopener noreferrer');
      openBtn.setAttribute('target', '_blank');

      // Handle both URL shortcuts and regular files with window.open for proper handling
      openBtn.href = '#';
      openBtn.addEventListener('click', (ev) => {
        ev.preventDefault();
        if (isUrlShortcut) {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          (async (): Promise<void> => {
            try {
              const target = await this._fetchUrlTarget(item);
              if (target) {
                window.open(target, '_blank', 'noopener,noreferrer');
              } else {
                const fallback = this._constructOpenUrl(item);
                window.open(fallback, '_blank', 'noopener,noreferrer');
              }
            } catch (error) {
              console.error('Failed to open .url target:', error);
              const fallback = this._constructOpenUrl(item);
              window.open(fallback, '_blank', 'noopener,noreferrer');
            }
          })();
        } else {
          // For regular files, also use window.open to ensure proper handling
          const fileUrl = this._constructOpenUrl(item);
          window.open(fileUrl, '_blank', 'noopener,noreferrer');
        }
      });

      openBtn.innerHTML = `<span class="${styles.buttonText}">Open</span>`;
      actions.appendChild(openBtn);

      if (!isUrlShortcut) {
        const downloadBtn = document.createElement('a');
        downloadBtn.className = `${styles.actionButton} ${styles.downloadButton}`;
        downloadBtn.title = 'Download';
        downloadBtn.setAttribute('rel', 'noopener noreferrer');
        downloadBtn.setAttribute('target', '_blank');
        downloadBtn.href = this._escapeHtml(this._constructDownloadUrl(item));
        downloadBtn.innerHTML = `<span class="${styles.buttonText}">Download</span>`;
        actions.appendChild(downloadBtn);
      }

      fileItem.appendChild(actions);
      fileListEl.appendChild(fileItem);
    });

    if (section.hasMoreFiles) {
      const remainingFiles = section.allFiles.length - section.displayedFiles.length;
      const loadMoreButton = document.createElement("button");
      loadMoreButton.className = styles.loadMoreButton;
      loadMoreButton.innerHTML = `Load ${remainingFiles} More File${remainingFiles !== 1 ? 's' : ''}`;
      loadMoreButton.addEventListener("click", () => {
        this._loadMoreFilesForSection(sectionId);
      });
      fileListEl.appendChild(loadMoreButton);
    }
  }

  private _constructOpenUrl(item: IFileItem): string {
    const siteAbsoluteUrl = this.context.pageContext.web.absoluteUrl;
    
    if ((item.ServerRelativeUrl as string).startsWith('http')) {
      return item.ServerRelativeUrl;
    }
    
    const siteDomain = siteAbsoluteUrl.split('/').slice(0, 3).join('/');
    let filePath = item.ServerRelativeUrl;
    if (!(filePath as string).startsWith('/')) {
      filePath = '/' + filePath;
    }
    
    return siteDomain + filePath;
  }

  private _constructDownloadUrl(item: IFileItem): string {
    const fileUrl = this._constructOpenUrl(item);
    return `${this.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(fileUrl)}`;
  }

  private async _fetchUrlTarget(item: IFileItem): Promise<string> {
    try {
      let serverRelative = item.ServerRelativeUrl || item.FileRef || '';
      if (!(serverRelative as string).startsWith('/')) {
        serverRelative = '/' + serverRelative;
      }

      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelative)}')/$value`;
      const response = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (!response.ok) {
        return '';
      }

      const content = await response.text();

      const match = content.match(/URL=(.*)/i);
      if (match && match[1]) {
        return match[1].trim();
      }

      const hrefMatch = content.match(/href=["']([^"']+)["']/i);
      if (hrefMatch && hrefMatch[1]) {
        return hrefMatch[1].trim();
      }

      return '';
    } catch (err) {
      console.error('Error fetching .url target:', err);
      return '';
    }
  }

  private _getFileExtension(filename: string): string {
    return filename.split('.').pop()?.toLowerCase() || '';
  }

  private _getFileTypeDisplay(fileType: string): string {
    const typeMap: { [key: string]: string } = {
      'pdf': 'PDF', 'doc': 'DOC', 'docx': 'DOCX', 'txt': 'TXT',
      'xls': 'XLS', 'xlsx': 'XLSX', 'ppt': 'PPT', 'pptx': 'PPTX',
      'odt': 'ODT', 'url': 'URL'
    };
    
    return typeMap[fileType.toLowerCase()] || fileType.toUpperCase();
  }

  private _escapeHtml(unsafe: string): string {
    return unsafe
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (!this.properties.sections) {
      this.properties.sections = [];
    }
    if (!this.properties.pageTitle) {
      this.properties.pageTitle = 'Resources';
    }

    const groups: IPropertyPaneGroup[] = [];

    // Page settings group
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const pageFields: IPropertyPaneField<any>[] = [
      PropertyPaneTextField('pageTitle', {
        label: 'Page Title',
        description: 'Enter the main title for this page',
        value: this.properties.pageTitle
      }),
      PropertyPaneDropdown('mainLibraryName', {
        label: 'Document Library *',
        options: this.isLibraryOptionsLoading 
          ? [{ key: 'loading', text: 'Loading libraries...' }]
          : this.libraryOptions.length > 0 
            ? this.libraryOptions 
            : [{ key: 'none', text: 'No libraries found' }],
        selectedKey: this.properties.mainLibraryName,
        disabled: this.isLibraryOptionsLoading
      }),
      PropertyPaneDropdown('mainFolderName', {
        label: 'Folder (Optional)',
        options: this.properties.mainLibraryName && this.folderOptions[this.properties.mainLibraryName]
          ? this.folderOptions[this.properties.mainLibraryName]
          : [{ key: '', text: 'Root folder (default)' }],
        selectedKey: this.properties.mainFolderName || '',
        disabled: !this.properties.mainLibraryName
      })
    ];

    groups.push({
      groupName: 'Page Settings',
      groupFields: pageFields
    });

    // Help content group
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const helpFields: IPropertyPaneField<any>[] = [
      PropertyPaneTextField('helpImageUrl', {
        label: 'Image URL (Optional)',
        description: 'Example: /sites/mysite/SiteAssets/help.png'
      }),
      PropertyPaneTextField('helpArticleHtml', {
        label: 'Help Content',
        description: 'Use **text** for bold, * for bullets, 1. for numbers',
        multiline: true,
        rows: 8
      })
    ];

    groups.push({
      groupName: 'Help Guide',
      groupFields: helpFields
    });

    // Add Link section (only visible if user has Edit permission)
    const userCanEdit = this.context.pageContext.web.permissions.hasPermission(SPPermission.editListItems);
    if (userCanEdit) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const addLinkFields: IPropertyPaneField<any>[] = [
        PropertyPaneButton('addLinkButton', {
          text: '+ Add Link',
          onClick: () => this._showAddLinkModal(),
          buttonType: 1
        })
      ];

      groups.push({
        groupName: 'Add Link',
        groupFields: addLinkFields
      });
    }

    // Additional sections
    this.properties.sections.forEach((section, index) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sectionFields: IPropertyPaneField<any>[] = [
        PropertyPaneTextField(`sections_${index}_sectionTitle`, {
          label: 'Section Title',
          description: 'Enter a title for this section',
          value: section.sectionTitle
        }),
        PropertyPaneDropdown(`sections_${index}_libraryName`, {
          label: 'Document Library *',
          options: this.isLibraryOptionsLoading 
            ? [{ key: 'loading', text: 'Loading libraries...' }]
            : this.libraryOptions.length > 0 
              ? this.libraryOptions 
              : [{ key: 'none', text: 'No libraries found' }],
          selectedKey: section.libraryName,
          disabled: this.isLibraryOptionsLoading
        }),
        PropertyPaneDropdown(`sections_${index}_folderName`, {
          label: 'Folder (Optional)',
          options: section.libraryName && this.folderOptions[section.libraryName]
            ? this.folderOptions[section.libraryName]
            : [{ key: '', text: 'Root folder (default)' }],
          selectedKey: section.folderName || '',
          disabled: !section.libraryName
        })
      ];

      sectionFields.push(
        PropertyPaneButton(`removeSection_${section.id}`, {
          text: `Remove Section ${index + 1}`,
          onClick: () => this._handleRemoveSection(section.id),
          buttonType: 0
        })
      );

      groups.push({
        groupName: `Additional Section ${index + 1}`,
        groupFields: sectionFields
      });
    });

    // Add section button
    if (this.properties.sections.length < 3) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const addSectionFields: IPropertyPaneField<any>[] = [
        PropertyPaneButton('addSection', {
          text: '+ Add Section',
          onClick: this._addSection.bind(this),
          buttonType: 1
        })
      ];

      groups.push({
        groupName: 'Add More',
        groupFields: addSectionFields
      });
    }

    return {
      pages: [
        {
          header: {
            description: 'Configure your resource page and sections'
          },
          groups: groups
        }
      ]
    };
  }

  private _handleRemoveSection(sectionId: string): void {
    const confirmDelete = confirm('Delete this section?');
    if (confirmDelete) {
      this._removeSection(sectionId);
    }
  }

  private _addSection(): void {
    if (!this.properties.sections) {
      this.properties.sections = [];
    }

    if (this.properties.sections.length >= 3) {
      alert('Maximum of 3 additional sections allowed');
      return;
    }

    const newSection: IResourceSection = {
      id: `section-${Date.now()}`,
      sectionTitle: `Additional Resources ${this.properties.sections.length + 1}`,
      libraryName: '',
      folderName: ''
    };

    this.properties.sections.push(newSection);
    this.context.propertyPane.refresh();
  }

  private _removeSection(sectionId: string): void {
    if (!this.properties.sections) {
      this.properties.sections = [];
    }

    this.properties.sections = this.properties.sections.filter(s => s.id !== sectionId);
    delete this.sectionData[sectionId];
    this.context.propertyPane.refresh();
    this.render();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | string[], newValue: string | string[]): void {
    if (propertyPath === 'pageTitle') {
      const newValueStr = Array.isArray(newValue) ? newValue[0] : newValue;
      this.properties.pageTitle = newValueStr;
      this.render();
      return;
    }

    if (propertyPath === 'mainLibraryName') {
      const newValueStr = Array.isArray(newValue) ? newValue[0] : newValue;
      this.properties.mainLibraryName = newValueStr;
      this.properties.mainFolderName = '';
      if (newValueStr) {
        this._loadFolderOptions(newValueStr).catch(console.error);
      }
      this.render();
      return;
    }

    if (propertyPath === 'mainFolderName') {
      const newValueStr = Array.isArray(newValue) ? newValue[0] : newValue;
      this.properties.mainFolderName = newValueStr || '';
      if (this.sectionData.main) {
        this.sectionData.main.currentFolder = newValueStr || '';
      }
      this.render();
      return;
    }

    const sectionMatch = propertyPath.match(/sections_(\d+)_(\w+)/);
    
    if (sectionMatch) {
      const sectionIndex = parseInt(sectionMatch[1], 10);
      const fieldName = sectionMatch[2];

      if (!this.properties.sections || !this.properties.sections[sectionIndex]) {
        return;
      }

      const section = this.properties.sections[sectionIndex];
      const newValueStr = Array.isArray(newValue) ? newValue[0] : newValue;

      if (fieldName === 'sectionTitle') {
        section.sectionTitle = newValueStr;
        this.render();
      } else if (fieldName === 'libraryName' && newValueStr) {
        section.libraryName = newValueStr;
        section.folderName = '';
        this._loadFolderOptions(newValueStr).catch(console.error);
      } else if (fieldName === 'folderName') {
        section.folderName = newValueStr || '';
        if (this.sectionData[section.id]) {
          this.sectionData[section.id].currentFolder = newValueStr || '';
        }
        this.render();
      }
    }

    this.context.propertyPane.refresh();
  }
}










