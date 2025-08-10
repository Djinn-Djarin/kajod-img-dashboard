<script>
    import * as XLSX from "xlsx";
    import { jsPDF } from "jspdf";
    import html2canvas from "html2canvas";
    import { tick } from "svelte";

    let cadets = [];
    let imagesMap = {};
    let excelLoaded = false;
    let imagesLoaded = false;

    let itemsPerPage = 8;
    let currentPage = 1;
    let imgHeightCm = 6; // default height
    let nameFontSize = 14; // default font size in px

    // Derived
    $: readyToShow = excelLoaded && imagesLoaded;
    $: totalPages = Math.ceil(cadets.length / itemsPerPage);
    $: paginatedCadets = cadets.slice(
        (currentPage - 1) * itemsPerPage,
        currentPage * itemsPerPage,
    );

    // Helper: result → color
    function getResultColorRGB(result) {
        if (!result) return "rgb(200,200,200)";
        result = result.trim().toUpperCase();
        if (result === "P" || result === "PASS") return "rgb(0,200,0)";
        if (result === "F" || result === "FAIL") return "rgb(200,0,0)";
        if (result === "AB" || result === "ABSENT") return "rgb(255,215,0)";
        return "rgb(200,200,200)";
    }

    /* =======================
	   Excel Upload Handler
	======================= */
    function handleExcelUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            const workbook = XLSX.read(new Uint8Array(e.target.result), {
                type: "array",
            });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet);
            console.log(jsonData[0]); // debug

            cadets = jsonData.map((row) => {
                const regNoKey = Object.keys(row).find(
                    (k) => k.trim() === "Regimental No.",
                );
                const nameKey = Object.keys(row).find(
                    (k) => k.trim() === "Name of Cadet",
                );
                const resultKey = Object.keys(row).find(
                    (k) => k.trim().toUpperCase() === "RESULT",
                );

                return {
                    regNo: String(row[regNoKey] || "").trim(),
                    name: String(row[nameKey] || "").trim(),
                    result: String(row[resultKey] || "").trim(),
                };
            });

            excelLoaded = true;
        };
        reader.readAsArrayBuffer(file);
    }

    /* =======================
	   Image Folder Upload Handler
	======================= */
    function handleImagesUpload(event) {
        const files = Array.from(event.target.files);
        if (!files.length) return;

        imagesMap = {};
        files
            .filter((file) => file.type.startsWith("image/"))
            .forEach((file) => {
                const nameWithoutExt = file.name
                    .replace(/\.[^/.]+$/, "")
                    .trim();
                imagesMap[nameWithoutExt] = URL.createObjectURL(file);
            });

        imagesLoaded = true;
    }

    /* =======================
	   Pagination
	======================= */
    function changePage(newPage) {
        if (newPage >= 1 && newPage <= totalPages) {
            currentPage = newPage;
        }
    }

    /* =======================
	   PDF Generator
	======================= */
    async function downloadPDF() {
        if (!readyToShow) return;

        const pdf = new jsPDF({
            orientation: "landscape",
            unit: "mm",
            format: "legal",
        });

        for (let pageIndex = 1; pageIndex <= totalPages; pageIndex++) {
            currentPage = pageIndex;
            await tick();

            const preview = document.getElementById("preview-area");
            if (!preview) continue;

            const clone = preview.cloneNode(true);
            clone.style.position = "absolute";
            clone.style.left = "-9999px";
            document.body.appendChild(clone);

            clone.querySelectorAll("*").forEach((el) => {
                const style = window.getComputedStyle(el);
                const safeBg = style.backgroundColor.startsWith("oklch")
                    ? "rgb(255,255,255)"
                    : style.backgroundColor;
                const safeColor = style.color.startsWith("oklch")
                    ? "rgb(0,0,0)"
                    : style.color;
                const safeBorder = style.borderColor.startsWith("oklch")
                    ? "rgb(0,0,0)"
                    : style.borderColor;

                el.style.backgroundColor = safeBg;
                el.style.color = safeColor;
                el.style.borderColor = safeBorder;
            });

            const canvas = await html2canvas(clone, {
                scale: 1.5,
                useCORS: true,
            });
            document.body.removeChild(clone);

            const imgData = canvas.toDataURL("image/png");
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();

            if (pageIndex > 1) pdf.addPage();
            pdf.addImage(imgData, "PNG", 0, 0, pageWidth, pageHeight);
        }

        pdf.save("cadets.pdf");
    }
</script>

<div class="flex h-screen bg-gray-50">
    <!-- Sidebar -->
    <aside
        class="w-72 bg-white shadow-lg p-6 flex flex-col space-y-6 border-r border-gray-200"
    >
        <h1 class="text-2xl font-bold text-blue-700">Cadet Viewer</h1>

        <div>
            <label class="block font-medium mb-1">Upload Excel</label>
            <input
                type="file"
                accept=".xlsx,.xls"
                on:change={handleExcelUpload}
                class="block w-full text-sm"
            />
        </div>

        <div>
            <label class="block font-medium mb-1">Upload Image Folder</label>
            <input
                type="file"
                webkitdirectory
                multiple
                on:change={handleImagesUpload}
                class="block w-full text-sm"
            />
        </div>

        <div>
            <label class="block font-medium mb-1">Images per page</label>
            <input
                type="number"
                min="1"
                bind:value={itemsPerPage}
                class="w-24 border rounded p-1"
            />
        </div>

        <div>
            <label class="block font-medium mb-1">Image height (cm)</label>
            <input
                type="number"
                min="1"
                step="0.5"
                bind:value={imgHeightCm}
                class="w-24 border rounded p-1"
            />
        </div>

        <div>
            <label class="block font-medium mb-1">Name font size (px)</label>
            <input
                type="number"
                min="8"
                step="1"
                bind:value={nameFontSize}
                class="w-24 border rounded p-1"
            />
        </div>

        <button
            on:click={downloadPDF}
            class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded shadow-md disabled:bg-gray-300"
            disabled={!readyToShow}
        >
            Download PDF
        </button>
    </aside>

    <!-- Main content -->
    <main class="flex-1 p-6 overflow-auto flex flex-col items-center">
        {#if readyToShow}
            <div
                id="preview-area"
                class="bg-white shadow-md p-4 rounded-lg mb-4"
                style="width: 35.56cm; height: 21.59cm;"
            >
                <div class="grid grid-cols-4 gap-4">
                    {#each paginatedCadets as cadet}
                        <div
                            class="bg-gray-50 rounded-lg shadow p-2 flex flex-col items-center border border-gray-200"
                        >
                            {#if imagesMap[cadet.regNo]}
                                <!-- Replace inside your cadet card -->
                                <div class="relative">
                                    <img
                                        src={imagesMap[cadet.regNo]}
                                        alt={cadet.name}
                                        class="object-contain mb-2"
                                        style="height:{imgHeightCm}cm"
                                    />
                                    <!-- Result Box with inline color -->
                                    {#if cadet.result}
                                        <div
                                            class="absolute bottom-1 right-1 w-4 h-4 rounded"
                                            style="background-color: {getResultColorRGB(
                                                cadet.result,
                                            )};"
                                        ></div>
                                    {/if}
                                </div>
                            {:else}
                                <div
                                    class="flex items-center justify-center text-gray-400 italic"
                                    style="height:{imgHeightCm}cm"
                                >
                                    No image
                                </div>
                            {/if}
                            <p
                                class="font-semibold text-gray-800"
                                style="font-size: {nameFontSize}px"
                            >
                                {cadet.regNo}
                            </p>
                            <p
                                class="text-sm text-gray-600"
                                style="font-size: {nameFontSize}px"
                            >
                                {cadet.name}
                            </p>
                        </div>
                    {/each}
                </div>
            </div>

            <div class="flex justify-center items-center space-x-4">
                <button
                    on:click={() => changePage(currentPage - 1)}
                    class="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
                    disabled={currentPage === 1}>Prev</button
                >
                <span>Page {currentPage} of {totalPages}</span>
                <button
                    on:click={() => changePage(currentPage + 1)}
                    class="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
                    disabled={currentPage === totalPages}>Next</button
                >
            </div>
        {:else}
            <div class="text-gray-500 italic text-center mt-10">
                Please upload both the Excel file and the image folder.
            </div>
        {/if}
    </main>
</div>
