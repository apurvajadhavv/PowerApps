import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface FileItem { name: string; type: string; size: number; content: string }

const GROUPS = {
  images: [".jpg", ".jpeg", ".png", ".gif"],
  documents: [".pdf", ".docx", ".xlsx"],
  textcsv: [".txt", ".csv"],
  media: [".mp3", ".mp4"]
};

export class FileUploader implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private host!: HTMLDivElement;
  private notifyOutputChanged!: () => void;
  private context!: ComponentFramework.Context<IInputs>;
  private selectedFiles: FileItem[] = [];
  private filesOutput = "";
  private lastResetValue = false;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.notifyOutputChanged = notifyOutputChanged;
    this.host = container;
    this.tryHydrateFromInitial();
    this.render();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    const resetNow = context.parameters.Reset.raw ?? false;
    const wasResetTriggered = resetNow !== this.lastResetValue && resetNow === true;
    this.lastResetValue = resetNow;

    this.context = context;

    if (wasResetTriggered) {
      this.selectedFiles = [];
      this.filesOutput = "[]";
      this.notifyOutputChanged();
    }

    this.render();
  }

  public getOutputs(): IOutputs {
    return { FilesOutput: this.filesOutput };
  }

  public destroy(): void {
    void 0;
  }

  private render(): void {
    const p = this.context.parameters;
    this.host.innerHTML = "";

    const wrapper = document.createElement("div");
    wrapper.className = "file-uploader-wrapper";

    const btn = document.createElement("button");
    btn.className = "upload-btn";
    btn.textContent = p.ButtonLabel.raw || "Upload Files";
    btn.style.backgroundColor = p.ButtonBackgroundColor.raw || "#0078D4";
    btn.style.color = p.ButtonTextColor.raw || "#FFFFFF";
    btn.style.borderRadius = `${p.ButtonBorderRadius.raw ?? 6}px`;
    btn.onmouseover = () => { btn.style.backgroundColor = p.ButtonHoverColor.raw || "#005A9E"; };
    btn.onmouseout = () => { btn.style.backgroundColor = p.ButtonBackgroundColor.raw || "#0078D4"; };

    const fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.style.display = "none";
    fileInput.multiple = p.AllowMultiple.raw ?? true;
    fileInput.accept = this.buildAcceptString();

    btn.onclick = () => fileInput.click();

    fileInput.onchange = async (e: Event) => {
      const el = e.target as HTMLInputElement;
      if (!el.files || el.files.length === 0) return;

      this.showOverlay("Processing files…");

      const chosen = Array.from(el.files);
      let added = false;

      for (const f of chosen) {
        const typeOk = this.isTypeAllowed(f);
        const sizeOk = this.validateSize(f);
        if (!typeOk || !sizeOk) {
          this.showToast(this.buildErrorText(f, typeOk, sizeOk), "error");
          continue;
        }
        await this.readFile(f);
        added = true;
      }

      this.hideOverlay();

      if (added) {
        this.updateFilesOutput();
        this.render();
      }

      el.value = "";
    };

    wrapper.appendChild(btn);
    wrapper.appendChild(fileInput);

    if ((p.ShowFileCount.raw ?? true) && this.selectedFiles.length > 0) {
      const count = document.createElement("div");
      count.className = "file-count";
      count.style.color = p.FileListTextColor.raw || "#333333";
      count.textContent = `${this.selectedFiles.length} file${this.selectedFiles.length > 1 ? "s" : ""} selected`;
      wrapper.appendChild(count);
    }

    if ((p.ShowFileList.raw ?? true) && this.selectedFiles.length > 0) {
      const listContainer = document.createElement("div");
      listContainer.className = "file-list-container";

      const ul = document.createElement("ul");
      ul.className = "file-list";

      this.selectedFiles.forEach((file, index) => {
        const li = document.createElement("li");
        li.className = "file-item";

        const info = document.createElement("div");
        info.className = "file-info";
        info.style.color = p.FileListTextColor.raw || "#333333";

        let html = `<strong>${file.name}</strong>`;
        if (p.ShowFileDetails.raw ?? true) {
          const sizeMB = (file.size / 1024 / 1024).toFixed(2);
          html += `<br>${file.type || "unknown"} — ${sizeMB} MB`;
        }
        info.innerHTML = html;

        const removeBtn = document.createElement("span");
        removeBtn.className = "remove-btn";
        removeBtn.style.color = p.FileListRemoveColor.raw || "#D13438";
        removeBtn.textContent = "❌";
        removeBtn.onclick = async () => {
          this.showOverlay("Removing file…");
          this.selectedFiles.splice(index, 1);
          this.updateFilesOutput();
          await new Promise((resolve) => setTimeout(resolve, 250));
          this.hideOverlay();
          this.render();
        };

        li.appendChild(info);
        li.appendChild(removeBtn);
        ul.appendChild(li);
      });

      listContainer.appendChild(ul);
      wrapper.appendChild(listContainer);
    }

    this.host.appendChild(wrapper);
  }

  private showOverlay(message: string): void {
    let overlay = this.host.querySelector(".overlay") as HTMLDivElement | null;
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.className = "overlay";
      const inner = document.createElement("div");
      inner.className = "overlay-inner";
      const spinner = document.createElement("div");
      spinner.className = "spinner";
      const msg = document.createElement("div");
      msg.className = "overlay-message";
      inner.appendChild(spinner);
      inner.appendChild(msg);
      overlay.appendChild(inner);
      const wrapper = this.host.querySelector(".file-uploader-wrapper");
      if (wrapper) wrapper.appendChild(overlay);
    }
    const msg = overlay.querySelector(".overlay-message") as HTMLDivElement;
    msg.textContent = message;
    overlay.classList.add("show");
  }

  private hideOverlay(): void {
    const overlay = this.host.querySelector(".overlay") as HTMLDivElement | null;
    if (overlay) overlay.classList.remove("show");
  }

  private showToast(message: string, kind: "info" | "error"): void {
    const dur = this.context.parameters.ToastDurationMS.raw ?? 2500;
    const toast = document.createElement("div");
    toast.className = `toast ${kind}`;
    toast.textContent = message;
    this.host.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add("show"));
    setTimeout(() => {
      toast.classList.remove("show");
      setTimeout(() => {
        if (toast.parentElement) toast.parentElement.removeChild(toast);
      }, 200);
    }, dur);
  }

  private buildAcceptString(): string {
    const p = this.context.parameters;
    const tokens: string[] = [];
    if (p.IncludeImages.raw ?? true) tokens.push(...GROUPS.images);
    if (p.IncludeDocuments.raw ?? true) tokens.push(...GROUPS.documents);
    if (p.IncludeTextCsv.raw ?? false) tokens.push(...GROUPS.textcsv);
    if (p.IncludeMedia.raw ?? false) tokens.push(...GROUPS.media);
    const extra = (p.CustomAcceptedTypes.raw || "")
      .split(",")
      .map((s) => s.trim())
      .filter((s) => !!s);
    return [...tokens, ...extra].join(",");
  }

  private isTypeAllowed(file: File): boolean {
    const accept = this.buildAcceptString();
    if (!accept) return true;
    const tokens = accept.split(",").map((t) => t.trim().toLowerCase());
    const name = file.name.toLowerCase();
    const ext = name.includes(".") ? `.${name.split(".").pop()}` : "";
    const mime = (file.type || "").toLowerCase();
    return tokens.some((t) => {
      if (t.startsWith(".")) return ext === t;
      if (t.endsWith("/*")) return mime.startsWith(t.slice(0, -1));
      return mime === t;
    });
  }

  private validateSize(file: File): boolean {
    const maxMB = this.context.parameters.MaxFileSizeMB.raw || 5;
    const maxBytes = maxMB * 1024 * 1024;
    return file.size <= maxBytes;
  }

  private buildErrorText(file: File, typeOk: boolean, sizeOk: boolean): string {
    const base = this.context.parameters.ErrorMessageText.raw || "This file is not allowed or exceeds the size limit.";
    const reasons: string[] = [];
    if (!typeOk) reasons.push("type not allowed");
    if (!sizeOk) reasons.push("too large");
    const extra = reasons.length ? ` (${reasons.join(", ")})` : "";
    return `"${file.name}" ${base}${extra}`;
  }

  private async readFile(file: File): Promise<void> {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => {
        const base64 = (reader.result as string).split(",")[1];
        this.selectedFiles.push({
          name: file.name,
          type: file.type,
          size: file.size,
          content: base64
        });
        resolve();
      };
      reader.readAsDataURL(file);
    });
  }

  private updateFilesOutput(): void {
    this.filesOutput = JSON.stringify(this.selectedFiles);
    this.notifyOutputChanged();
  }

  private tryHydrateFromInitial(): void {
    const initial = this.context.parameters.FilesInitial.raw;
    if (!initial) return;
    try {
      const arr = JSON.parse(initial) as FileItem[];
      this.selectedFiles = Array.isArray(arr) ? arr : [];
      this.filesOutput = JSON.stringify(this.selectedFiles);
      this.notifyOutputChanged();
    } catch (err) {
      void err;
    }
  }
}