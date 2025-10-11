//
// **图片处理与文档生成系统 v2**
//
// 实现了现代化的UI和工作流程:
// 1. **选项卡式界面**: 清晰地分为“选择与合并”和“预览与导出”两个步骤。
// 2. **引导式操作**: 用户完成第一步后，系统会自动切换到第二步，且相关按钮会随之解锁。
// 3. **模态框进度**: 所有耗时操作（合并、打包、生成Excel）都会在模态框中显示进度，不阻塞主界面。
// 4. **文件选择与分组**: 核心逻辑不变，根据文件名自动分组。
// 5. **图片合并**: 2x2网格合并功能保持不变。
// 6. **结果处理**: 预览、打包下载和导出Excel功能被整合到第二步，界面更整洁。
//

declare const ExcelJS: any;
declare const JSZip: any;

interface MergedImage {
    id: string;
    blob: Blob;
    url: string;
    fileName: string;
}

// UI Element References
const imageInput = document.getElementById('image-input') as HTMLInputElement;
const statusInfo = document.getElementById('status-info') as HTMLDivElement;
const mergeButton = document.getElementById('merge-button') as HTMLButtonElement;
const exportExcelButton = document.getElementById('export-excel-button') as HTMLButtonElement;
const downloadZipButton = document.getElementById('download-zip-button') as HTMLButtonElement;
const imageGrid = document.getElementById('image-grid') as HTMLDivElement;

// Modal References
const progressModal = document.getElementById('progress-modal') as HTMLElement;
const progressLabel = document.getElementById('progress-label') as HTMLHeadingElement;
const progressBar = document.getElementById('progress-bar') as HTMLDivElement;
const progressText = document.getElementById('progress-text') as HTMLDivElement;

// Tab References
const tabButtons = document.querySelectorAll('.tab-button');
const tabPanels = document.querySelectorAll('.tab-panel');
const tabStep2 = document.getElementById('tab-step2') as HTMLButtonElement;


let imageGroups = new Map<string, File[]>();
let mergedImages: MergedImage[] = [];

// --- Event Listeners ---

/**
 * 当用户选择文件时触发
 */
imageInput.addEventListener('change', () => {
    const files = imageInput.files;
    if (!files || files.length === 0) {
        return;
    }
    
    resetState();

    for (const file of Array.from(files)) {
        const prefix = file.name.split('-')[0];
        if (!imageGroups.has(prefix)) {
            imageGroups.set(prefix, []);
        }
        imageGroups.get(prefix)!.push(file);
    }
    
    const validGroupsCount = Array.from(imageGroups.values()).filter(group => group.length === 4).length;

    statusInfo.textContent = `已选择 ${files.length} 个文件，识别出 ${imageGroups.size} 个分组，其中 ${validGroupsCount} 个分组有效（包含4张图片）。`;

    if (validGroupsCount > 0) {
        mergeButton.disabled = false;
    } else {
        statusInfo.textContent += ' 没有找到有效的分组，请检查文件命名。';
        mergeButton.disabled = true;
    }
});

/**
 * 点击“开始合并”按钮时触发
 */
mergeButton.addEventListener('click', async () => {
    const validGroups = Array.from(imageGroups.entries()).filter(([_, files]) => files.length === 4);
    if (validGroups.length === 0) return;

    mergeButton.disabled = true;
    showModal('正在合并图片...');
    
    let processedCount = 0;
    const totalCount = validGroups.length;

    for (const [prefix, files] of validGroups) {
        try {
            const mergedBlob = await mergeGroupOfFour(files);
            const fileName = `${prefix}-merged.jpg`;
            const url = URL.createObjectURL(mergedBlob);
            mergedImages.push({ id: prefix, blob: mergedBlob, url, fileName });
        } catch (error) {
            console.error(`处理分组 ${prefix} 时出错:`, error);
        }
        processedCount++;
        updateProgress(processedCount, totalCount);
    }

    if (mergedImages.length > 0) {
        displayAllMergedImages();
        exportExcelButton.disabled = false;
        downloadZipButton.disabled = false;
        tabStep2.disabled = false;
        switchTab('step2');
    }
    
    hideModal();
});

/**
 * 点击“打包下载”按钮时触发
 */
downloadZipButton.addEventListener('click', async () => {
    if (mergedImages.length === 0) return;
    
    showModal('正在打包ZIP文件...');
    updateProgress(0, mergedImages.length);
    
    const zip = new JSZip();
    let processedCount = 0;
    for (const image of mergedImages) {
        const zipFileName = `${image.id}.jpg`;
        zip.file(zipFileName, image.blob);
        processedCount++;
        updateProgress(processedCount, mergedImages.length);
    }
    
    const content = await zip.generateAsync({type:"blob"});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(content);
    link.download = "merged-images.zip";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    hideModal();
});

/**
 * 点击“导出Excel”按钮时触发
 */
exportExcelButton.addEventListener('click', async () => {
    if (mergedImages.length === 0) return;

    showModal('正在生成 Excel 文件...');
    updateProgress(0, mergedImages.length);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('图片处理报告');

    worksheet.columns = [
        { header: '图片编号', key: 'id', width: 30 },
        { header: '图片预览', key: 'image', width: 40 },
        { header: '原始图片数量', key: 'count', width: 20 },
        { header: '处理状态', key: 'status', width: 20 },
    ];
    
    let processedCount = 0;
    for (const image of mergedImages) {
        const row = worksheet.addRow({ id: image.id, count: 4, status: '已合并' });

        const base64 = await blobToBase64(image.blob);

        const imageId = workbook.addImage({ base64: base64.split(',')[1], extension: 'jpeg' });
        
        worksheet.getRow(row.number).height = 160;
        worksheet.addImage(imageId, {
            tl: { col: 1.1, row: row.number - 0.9 },
            ext: { width: 200, height: 200 },
            editAs: 'oneCell'
        });
        
        processedCount++;
        updateProgress(processedCount, mergedImages.length);
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '图片处理报告.xlsx';
    a.click();
    URL.revokeObjectURL(url);
    
    hideModal();
});

// --- Core Logic ---

/**
 * 将一组4张图片合并为一张
 */
function mergeGroupOfFour(files: File[]): Promise<Blob> {
    return new Promise((resolve, reject) => {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        if (!ctx) return reject(new Error('无法获取Canvas上下文'));

        files.sort((a, b) => {
            const numA = parseInt(a.name.match(/-(\d+)\./)?.[1] || '0');
            const numB = parseInt(b.name.match(/-(\d+)\./)?.[1] || '0');
            return numA - numB;
        });

        const imagePromises = files.map(file => new Promise<HTMLImageElement>((resolveImg, rejectImg) => {
            const img = new Image();
            img.onload = () => resolveImg(img);
            img.onerror = rejectImg;
            img.src = URL.createObjectURL(file);
        }));

        Promise.all(imagePromises).then(images => {
            const [img1, img2, img3, img4] = images;
            const width = Math.max(...images.map(img => img.width));
            const height = Math.max(...images.map(img => img.height));
            
            canvas.width = width * 2;
            canvas.height = height * 2;

            ctx.drawImage(img1, 0, 0, width, height);
            ctx.drawImage(img2, width, 0, width, height);
            ctx.drawImage(img3, 0, height, width, height);
            ctx.drawImage(img4, width, height, width, height);

            images.forEach(img => URL.revokeObjectURL(img.src));

            canvas.toBlob(blob => {
                if (blob) resolve(blob);
                else reject(new Error('Canvas toBlob失败'));
            }, 'image/jpeg', 0.95);
        }).catch(reject);
    });
}

/**
 * 将Blob对象转换为Base64字符串
 */
function blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

// --- UI Management ---

/**
 * 显示所有合并后的图片
 */
function displayAllMergedImages() {
    imageGrid.innerHTML = '';
    mergedImages.forEach(mergedImage => {
        const card = document.createElement('div');
        card.className = 'image-card';
        card.innerHTML = `
            <img src="${mergedImage.url}" alt="合并后的图片 ${mergedImage.id}">
            <div class="image-card-info">
                <p>${mergedImage.id}</p>
                <a href="${mergedImage.url}" download="${mergedImage.fileName}" class="button primary-button">下载</a>
            </div>
        `;
        imageGrid.appendChild(card);
    });
}

/**
 * 更新进度条
 */
function updateProgress(current: number, total: number) {
    const percent = total > 0 ? (current / total) * 100 : 0;
    progressBar.style.width = `${percent}%`;
    progressText.textContent = `${current} / ${total}`;
}

/**
 * 显示进度模态框
 */
function showModal(label: string) {
    progressLabel.textContent = label;
    updateProgress(0, 0);
    progressModal.style.display = 'flex';
}

/**
 * 隐藏进度模态框
 */
function hideModal() {
    setTimeout(() => {
        progressModal.style.display = 'none';
    }, 500); // Delay to show completion
}

/**
 * 重置应用状态
 */
function resetState() {
    imageGroups.clear();
    mergedImages.forEach(img => URL.revokeObjectURL(img.url));
    mergedImages = [];
    imageGrid.innerHTML = '<p class="placeholder">合并后的图片将显示在这里。</p>';
    statusInfo.textContent = '';
    mergeButton.disabled = true;
    exportExcelButton.disabled = true;
    downloadZipButton.disabled = true;
    tabStep2.disabled = true;
    switchTab('step1');
}

/**
 * 切换选项卡
 */
function switchTab(tabId: string) {
    tabPanels.forEach(panel => {
        panel.classList.toggle('active', panel.id === `panel-${tabId}`);
    });
    tabButtons.forEach(button => {
        button.classList.toggle('active', (button as HTMLElement).dataset.tab === tabId);
    });
}

tabButtons.forEach(button => {
    button.addEventListener('click', () => {
        if (!button.hasAttribute('disabled')) {
            switchTab((button as HTMLElement).dataset.tab!);
        }
    });
});
