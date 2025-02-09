document.addEventListener("DOMContentLoaded", function () {
  const numStudentsInput = document.getElementById("numStudents");
  const rowsInput = document.getElementById("rows");
  const colsInput = document.getElementById("cols");
  const generateButton = document.getElementById("generateSeats");
  const randomizeButton = document.getElementById("randomize");
  const seatingChart = document.getElementById("seatingChart");
  const printBtn = document.getElementById("printBtn");
  const fileInput = document.getElementById("fileInput");
  const importBtn = document.getElementById("importBtn");
  const fileName = document.getElementById("fileName");
  const columnSelectModal = document.getElementById("columnSelectModal");
  const columnList = document.getElementById("columnList");
  const confirmColumnBtn = document.getElementById("confirmColumn");
  const cancelColumnBtn = document.getElementById("cancelColumn");

  let seats = [];
  let dragSrcEl = null;
  let importedNames = [];
  let excelData = null;
  let selectedCells = new Set(); // 存储选中的单元格
  let isMouseDown = false; // 跟踪鼠标按下状态
  let isSelecting = null; // 跟踪是选中还是取消选中的操作

  // 加载保存的座位信息
  function loadSavedSeats() {
    const savedData = localStorage.getItem("seatingData");
    if (savedData) {
      const data = JSON.parse(savedData);
      numStudentsInput.value = data.numStudents;
      rowsInput.value = data.rows;
      colsInput.value = data.cols;
      importedNames = data.names;

      // 自动生成座位
      generateSeats();
    }
  }

  // 保存座位信息
  function saveSeatingData() {
    const seatingData = {
      numStudents: numStudentsInput.value,
      rows: rowsInput.value,
      cols: colsInput.value,
      names: Array.from(
        document.querySelectorAll(".seat:not(.empty) input")
      ).map((input) => input.value),
    };
    localStorage.setItem("seatingData", JSON.stringify(seatingData));
  }

  // 在页面加载时恢复座位
  loadSavedSeats();

  // 绑定生成座位按钮的点击事件
  generateButton.addEventListener("click", generateSeats);

  function handleDragStart(e) {
    this.style.opacity = "0.4";
    dragSrcEl = this;
    e.dataTransfer.effectAllowed = "move";
  }

  function handleDragEnd(e) {
    this.style.opacity = "1";
    document.querySelectorAll(".seat").forEach((seat) => {
      seat.classList.remove("drag-over");
    });
  }

  function handleDragOver(e) {
    e.preventDefault();
    return false;
  }

  function handleDragEnter(e) {
    this.classList.add("drag-over");
  }

  function handleDragLeave(e) {
    this.classList.remove("drag-over");
  }

  function handleDrop(e) {
    e.stopPropagation();

    if (dragSrcEl !== this) {
      // 交换两个座位的输入值和状态
      const srcInput = dragSrcEl.querySelector("input");
      const destInput = this.querySelector("input");

      // 如果目标是空座位
      if (this.classList.contains("empty")) {
        // 移动到空座位
        destInput.value = srcInput.value;
        this.dataset.value = srcInput.value;
        srcInput.value = "";
        dragSrcEl.dataset.value = "";
        destInput.disabled = false;
        srcInput.disabled = true;

        // 更新座位状态
        this.classList.remove("empty");
        dragSrcEl.classList.add("empty");

        // 更新占位符
        destInput.placeholder = `座位 ${
          Array.from(seatingChart.children).indexOf(this) + 1
        }`;
        srcInput.placeholder = "空座";

        // 交换两个座位的拖拽状态
        const tempListeners = {
          dragstart: dragSrcEl._dragstart,
          dragend: dragSrcEl._dragend,
          dragover: dragSrcEl._dragover,
          dragenter: dragSrcEl._dragenter,
          dragleave: dragSrcEl._dragleave,
          drop: dragSrcEl._drop,
        };

        // 保存事件监听器引用
        dragSrcEl._dragstart = handleDragStart;
        dragSrcEl._dragend = handleDragEnd;
        dragSrcEl._dragover = handleDragOver;
        dragSrcEl._dragenter = handleDragEnter;
        dragSrcEl._dragleave = handleDragLeave;
        dragSrcEl._drop = handleDrop;

        this._dragstart = tempListeners.dragstart;
        this._dragend = tempListeners.dragend;
        this._dragover = tempListeners.dragover;
        this._dragenter = tempListeners.dragenter;
        this._dragleave = tempListeners.dragleave;
        this._drop = tempListeners.drop;

        // 更新seats数组
        const srcIndex = seats.indexOf(dragSrcEl);
        if (srcIndex > -1) {
          seats.splice(srcIndex, 1);
        }
        seats.push(this);
      } else {
        // 普通座位之间交换
        const tempValue = srcInput.value;
        srcInput.value = destInput.value;
        destInput.value = tempValue;
        // 更新 data-value
        const tempDataValue = dragSrcEl.dataset.value;
        dragSrcEl.dataset.value = this.dataset.value;
        this.dataset.value = tempDataValue;
      }
    }

    return false;
  }

  function removeDragListeners(seat) {
    seat.removeAttribute("draggable");
    seat.removeEventListener("dragstart", handleDragStart);
    seat.removeEventListener("dragend", handleDragEnd);
    seat.removeEventListener("dragover", handleDragOver);
    seat.removeEventListener("dragenter", handleDragEnter);
    seat.removeEventListener("dragleave", handleDragLeave);
    seat.removeEventListener("drop", handleDrop);
  }

  function addDragListeners(seat) {
    seat.setAttribute("draggable", true);
    seat.addEventListener("dragstart", handleDragStart);
    seat.addEventListener("dragend", handleDragEnd);
    seat.addEventListener("dragover", handleDragOver);
    seat.addEventListener("dragenter", handleDragEnter);
    seat.addEventListener("dragleave", handleDragLeave);
    seat.addEventListener("drop", handleDrop);
  }

  // 导入按钮点击事件
  importBtn.addEventListener("click", function () {
    // 创建一个新的文件输入元素
    const newFileInput = document.createElement("input");
    newFileInput.type = "file";
    newFileInput.accept = ".xlsx,.xls,.csv";
    newFileInput.style.display = "none";

    // 复制原始文件输入的事件处理程序
    newFileInput.addEventListener("change", function (e) {
      const file = e.target.files[0];
      if (!file) return;

      fileName.textContent = file.name;
      const reader = new FileReader();

      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

          if (excelData && excelData.length > 0) {
            const columnCount = excelData[0].length;
            showColumnSelectModal(new Array(columnCount).fill(null));
            selectedCells.clear();
          } else {
            alert("文件中没有数据");
          }
        } catch (error) {
          console.error("Error reading Excel:", error);
          alert("文件读取失败，请确保文件格式正确");
        }
      };

      reader.readAsArrayBuffer(file);

      // 移除临时创建的输入元素
      document.body.removeChild(newFileInput);
    });

    // 添加到文档并触发点击
    document.body.appendChild(newFileInput);
    newFileInput.click();
  });

  // 显示列选择弹窗
  function showColumnSelectModal(columns) {
    columnList.innerHTML = "";
    // 创建表格视图
    const table = document.createElement("table");
    table.className = "excel-table";

    // 创建表格内容
    const tbody = document.createElement("tbody");
    excelData.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");

      row.forEach((cell, colIndex) => {
        const td = document.createElement("td");
        if (cell) {
          td.textContent = cell;
          td.dataset.value = cell;

          // 鼠标按下事件
          td.addEventListener("mousedown", (e) => {
            isMouseDown = true;
            isSelecting = !td.classList.contains("selected");
            toggleCell(td);
            e.preventDefault(); // 防止文本选择
          });

          // 鼠标进入事件
          td.addEventListener("mouseenter", () => {
            if (isMouseDown) {
              toggleCell(td);
            }
          });
        }
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    columnList.appendChild(table);

    // 添加鼠标抬起事件
    document.addEventListener("mouseup", () => {
      isMouseDown = false;
    });

    // 防止在拖动选择时选中文本
    table.addEventListener("selectstart", (e) => e.preventDefault());

    // 初始化确认按钮状态
    confirmColumnBtn.disabled = true;
    columnSelectModal.classList.remove("hidden");
  }

  // 切换单元格选中状态
  function toggleCell(td) {
    const cell = td.dataset.value;
    if (isSelecting) {
      td.classList.add("selected");
      selectedCells.add(cell);
    } else {
      td.classList.remove("selected");
      selectedCells.delete(cell);
    }
    confirmColumnBtn.disabled = selectedCells.size === 0;
  }

  // 隐藏列选择弹窗
  function hideColumnSelectModal() {
    columnSelectModal.classList.add("hidden");
    selectedCells.clear(); // 清空选中状态
  }

  // 检查重复名字
  function findDuplicateNames(names) {
    const duplicates = {};
    names.forEach((name, index) => {
      if (name.trim()) {
        if (!duplicates[name]) {
          duplicates[name] = [index];
        } else {
          duplicates[name].push(index);
        }
      }
    });
    return Object.entries(duplicates).filter(
      ([_, indices]) => indices.length > 1
    );
  }

  // 显示重复名字处理弹窗
  function showDuplicateNamesModal(duplicates, names, callback) {
    const modal = document.getElementById("duplicateNamesModal");
    const duplicateList = document.getElementById("duplicateList");
    duplicateList.innerHTML = "";

    duplicates.forEach(([name, indices]) => {
      const group = document.createElement("div");
      group.className = "duplicate-group";
      group.innerHTML = `<div class="duplicate-name">「${name}」出现了 ${indices.length} 次：</div>`;

      indices.forEach((index) => {
        const item = document.createElement("label");
        item.className = "duplicate-item";
        item.innerHTML = `
          <input type="checkbox" value="${index}" checked>
          第 ${index + 1} 个位置
        `;
        group.appendChild(item);
      });

      duplicateList.appendChild(group);
    });

    const confirmBtn = document.getElementById("confirmDuplicate");
    const cancelBtn = document.getElementById("cancelDuplicate");

    // 确认按钮事件
    const handleConfirm = () => {
      const selectedIndices = new Set(
        Array.from(duplicateList.querySelectorAll("input:checked")).map(
          (input) => parseInt(input.value)
        )
      );

      // 过滤掉未选中的重复项
      const filteredNames = names.filter((_, index) =>
        selectedIndices.has(index)
      );

      modal.classList.add("hidden");
      confirmBtn.removeEventListener("click", handleConfirm);
      cancelBtn.removeEventListener("click", handleCancel);

      callback(filteredNames);
    };

    // 取消按钮事件
    const handleCancel = () => {
      modal.classList.add("hidden");
      confirmBtn.removeEventListener("click", handleConfirm);
      cancelBtn.removeEventListener("click", handleCancel);
    };

    confirmBtn.addEventListener("click", handleConfirm);
    cancelBtn.addEventListener("click", handleCancel);
    modal.classList.remove("hidden");
  }

  // 确认按钮点击事件
  confirmColumnBtn.addEventListener("click", () => {
    if (selectedCells.size === 0) {
      alert("请选择一列数据");
      return;
    }

    // 获取当前已有的名字
    const existingNames = [];
    document.querySelectorAll(".seat:not(.empty) input").forEach((input) => {
      if (input.value.trim()) {
        existingNames.push(input.value.trim());
      }
    });

    // 获取所有选中的值
    const newNames = Array.from(selectedCells);

    // 合并现有名字和新导入的名字
    const allNames = [...newNames, ...existingNames];

    // 检查重复名字
    const duplicates = findDuplicateNames(allNames);
    if (duplicates.length > 0) {
      showDuplicateNamesModal(duplicates, allNames, (filteredNames) => {
        importedNames = filteredNames;
        processImportedNames();
      });
    } else {
      importedNames = allNames;
      processImportedNames();
    }

    hideColumnSelectModal();

    // 保存当前状态
    saveSeatingData();
  });

  // 取消按钮点击事件
  cancelColumnBtn.addEventListener("click", hideColumnSelectModal);

  // 当输入人数时自动计算推荐的行列数
  numStudentsInput.addEventListener("input", function () {
    const numStudents = parseInt(this.value) || 0;
    if (numStudents > 0) {
      const sqrt = Math.sqrt(numStudents);
      const recommendedCols = Math.ceil(sqrt);
      const recommendedRows = Math.ceil(numStudents / recommendedCols);

      rowsInput.value = recommendedRows;
      colsInput.value = recommendedCols;
    }
  });

  // 修改生成座位的函数
  function generateSeats() {
    const numStudents = parseInt(numStudentsInput.value);
    const rows = parseInt(rowsInput.value) || 0;
    const cols = parseInt(colsInput.value) || 0;

    if (numStudents < 1 || isNaN(numStudents)) {
      alert("请输入有效的人数！");
      return;
    }

    if (rows < 1 || cols < 1) {
      alert("请输入有效的行列数！");
      return;
    }

    if (rows * cols < numStudents) {
      alert("请输入有效的行列数，且确保座位数足够容纳所有学生！");
      return;
    }

    // 显示整个教室布局
    document.getElementById("classroom-layout").classList.remove("hidden");

    // 设置网格布局
    seatingChart.style.gridTemplateColumns = `repeat(${cols}, 100px)`;

    // 清空现有座位
    seatingChart.innerHTML = "";
    seats = [];

    // 生成座位
    for (let i = 0; i < rows * cols; i++) {
      const seat = document.createElement("div");
      seat.className = i < numStudents ? "seat" : "seat empty";

      const input = document.createElement("input");
      input.type = "text";
      input.placeholder = `座位 ${i + 1}`;
      if (i < numStudents && importedNames[i]) {
        input.value = importedNames[i];
        seat.dataset.value = importedNames[i];
      }
      // 添加输入事件监听器
      input.addEventListener("input", function () {
        seat.dataset.value = this.value;
      });

      if (i >= numStudents) {
        input.disabled = true;
        input.placeholder = "空座";
        // 为空座位添加完整的拖拽监听器
        seat.setAttribute("draggable", true);
        seat.addEventListener("dragstart", handleDragStart);
        seat.addEventListener("dragend", handleDragEnd);
        seat.addEventListener("dragover", handleDragOver);
        seat.addEventListener("dragenter", handleDragEnter);
        seat.addEventListener("dragleave", handleDragLeave);
        seat.addEventListener("drop", handleDrop);
      }

      seat.appendChild(input);
      seatingChart.appendChild(seat);
      if (i < numStudents) {
        seats.push(seat);
        addDragListeners(seat);
      }
    }

    // 显示随机调整按钮
    randomizeButton.classList.remove("hidden");

    // 在生成座位后保存数据
    saveSeatingData();
  }

  // 在输入框值改变时保存数据
  document.querySelectorAll(".seat input").forEach((input) => {
    input.addEventListener("change", saveSeatingData);
  });

  // 在随机调整座位后保存数据
  randomizeButton.addEventListener("click", function () {
    // 获取所有姓名
    const names = seats.map((seat) => seat.querySelector("input").value);

    // 打乱数组
    for (let i = names.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [names[i], names[j]] = [names[j], names[i]];
    }

    // 应用新的座位安排后保存数据
    seats.forEach((seat, index) => {
      const input = seat.querySelector("input");
      input.style.opacity = "0";
      setTimeout(() => {
        input.value = names[index];
        seat.dataset.value = names[index];
        input.style.opacity = "1";
        saveSeatingData();
      }, 300);
    });
  });

  // 打印功能
  printBtn.addEventListener("click", function () {
    const layout = document.getElementById("classroom-layout");
    if (layout.classList.contains("hidden")) {
      alert("请先生成座位表");
      return;
    }

    // 检查是否有空座位
    const emptySeats = document.querySelectorAll(".seat.empty");
    if (emptySeats.length > 0) {
      // 显示打印选项对话框
      document.getElementById("printOptionModal").classList.remove("hidden");
    } else {
      // 没有空座位，直接打印
      window.print();
    }
  });

  // 确认打印按钮事件
  document
    .getElementById("confirmPrint")
    .addEventListener("click", function () {
      const printEmptySeats =
        document.getElementById("printEmptySeats").checked;

      // 临时隐藏不需要打印的空座位
      const emptySeats = document.querySelectorAll(".seat.empty");
      if (!printEmptySeats) {
        emptySeats.forEach((seat) => (seat.style.display = "none"));
      }

      // 隐藏打印选项对话框
      document.getElementById("printOptionModal").classList.add("hidden");

      // 打印
      window.print();

      // 恢复空座位显示
      if (!printEmptySeats) {
        setTimeout(() => {
          emptySeats.forEach((seat) => (seat.style.display = ""));
        }, 100);
      }
    });

  // 取消打印按钮事件
  document.getElementById("cancelPrint").addEventListener("click", function () {
    document.getElementById("printOptionModal").classList.add("hidden");
  });

  // 添加清除数据的功能（可选）
  window.clearSeatingData = function () {
    localStorage.removeItem("seatingData");
    location.reload();
  };

  // 重新开始功能
  const restartBtn = document.getElementById("restartBtn");
  restartBtn.addEventListener("click", function () {
    if (confirm("确定要重新开始吗？这将清空所有座位信息。")) {
      // 清空所有输入
      numStudentsInput.value = "";
      rowsInput.value = "";
      colsInput.value = "";
      importedNames = [];

      // 隐藏座位表和随机按钮
      document.getElementById("classroom-layout").classList.add("hidden");
      randomizeButton.classList.add("hidden");

      // 清空本地存储
      localStorage.removeItem("seatingData");

      // 清空文件选择
      fileInput.value = "";
      fileName.textContent = "";
    }
  });

  // 将原有的处理逻辑移到单独的函数中
  function processImportedNames() {
    // 如果已经生成了座位表，直接更新现有座位
    const layout = document.getElementById("classroom-layout");
    if (!layout.classList.contains("hidden")) {
      const allInputs = document.querySelectorAll(".seat:not(.empty) input");
      const currentSeats = allInputs.length;

      // 如果导入的名字数量超过当前座位数，自动增加座位
      if (importedNames.length > currentSeats) {
        // 保存当前的行列数
        const currentRows = parseInt(rowsInput.value);
        const currentCols = parseInt(colsInput.value);

        // 计算需要的新行数
        const neededSeats = importedNames.length;
        const newRows = Math.ceil(neededSeats / currentCols);

        // 更新行数和人数
        rowsInput.value = newRows;
        numStudentsInput.value = neededSeats;

        // 重新生成座位表
        generateSeats();

        alert(
          `座位数不足，已自动增加行数到 ${newRows} 行，以容纳所有 ${neededSeats} 个名字`
        );
      } else {
        // 只更新现有座位的内容，不改变座位数量
        allInputs.forEach((input, index) => {
          if (index < importedNames.length) {
            input.value = importedNames[index];
            input.parentElement.dataset.value = importedNames[index];
          }
        });

        alert(`成功导入 ${importedNames.length} 个名字`);
      }
    } else {
      // 自动设置人数和推荐的行列数
      const numStudents = importedNames.length;
      numStudentsInput.value = numStudents;

      // 计算推荐的行列数
      const sqrt = Math.sqrt(numStudents);
      const recommendedCols = Math.ceil(sqrt);
      const recommendedRows = Math.ceil(numStudents / recommendedCols);

      rowsInput.value = recommendedRows;
      colsInput.value = recommendedCols;

      // 生成座位表
      generateSeats();

      alert(`成功导入 ${importedNames.length} 个名字并生成座位表`);
    }
  }
});
