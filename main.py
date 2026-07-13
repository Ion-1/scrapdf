from reportlab.lib.utils import ImageReader
import io

import os
import sys
import time
import shutil
import atexit
import datetime
import dataclasses

import polars as pl
import pikepdf as pike
import qtawesome as qta
import playwright.sync_api as psa
import playwright.async_api as paa
import PySide6.QtAsyncio as QtAsyncio

from pathlib import Path
from typing import cast, overload, TypedDict

from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from rust_enum import enum, Case, Result
from PySide6.QtCore import (
    Qt,
    QAbstractTableModel,
    QModelIndex,
    Slot,
    QSize,
    Signal,
    QRect,
    QObject,
    QThread, QSortFilterProxyModel,
)
from PySide6.QtGui import (
    QShortcut,
    QKeySequence,
    QPixmap,
    QIcon,
    QIconEngine,
    QPainter,
    QImageReader,
    QColor,
    QBrush,
    QPen,
    QCursor,
)
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QFileDialog,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QStackedWidget,
    QErrorMessage,
    QTabWidget,
    QTableView,
    QListWidget,
    QHBoxLayout,
    QHeaderView,
    QAbstractItemView,
    QInputDialog,
    QMessageBox,
    QProgressBar,
    QScrollArea,
    QListWidgetItem,
    QSizePolicy,
    QStyle,
    QDialog,
    QFrame,
    QCheckBox,
    QProxyStyle,
    QStyle,
)

os.environ["QT_DEBUG_LAYOUT"] = "1"
LOCAL_DIR = Path(os.path.expandvars("%LOCALAPPDATA%")) / Path("ScraPDF")
LOCAL_DIR.mkdir(parents=False, exist_ok=True)

BIN_DATA_DIR = Path(__file__).parent / Path("data")
LOCAL_DATA_DIR = LOCAL_DIR / Path("data")
if not LOCAL_DATA_DIR.exists():
    replacement_files = [
        Path("firefox") / Path("policies.json"),
    ]
    BIN_DATA_DIR.copy_into(LOCAL_DIR)
    for replacement_file in replacement_files:
        replacement_file = LOCAL_DATA_DIR / replacement_file
        replace = {"%LOCAL_DATA_DIR_URI%": LOCAL_DATA_DIR.as_uri()}
        file_bytes = replacement_file.read_bytes()
        for old, new in replace.items():
            file_bytes = file_bytes.replace(old.encode(), new.encode())
        replacement_file.write_bytes(file_bytes)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.stack.insertWidget(0, FileSelectionWidget())

        self.resize(1200, 800)


class FileSelectionWidget(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout(self)

        self.text = QLabel("Select a file")
        layout.addWidget(self.text, alignment=Qt.AlignmentFlag.AlignCenter)

        self.button = QPushButton("Browse")
        layout.addWidget(self.button, alignment=Qt.AlignmentFlag.AlignCenter)
        self.button.clicked.connect(self._on_browse_clicked)

    def _on_browse_clicked(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select a spreadsheet",
            "",
            "Spreadsheets (*.xlsx *.xls *.csv *.ods);;All Files (*.*)",
        )

        if not filename:
            return

        self.filename = filename
        self.path_fname = Path(filename)

        try:
            if self.path_fname.suffix in (".xlsx", ".xls"):
                df = pl.read_excel(filename, has_header=False, sheet_id=0)
            elif self.path_fname == ".csv":
                df = pl.read_csv(filename, has_header=False)
            elif self.path_fname == ".ods":
                df = pl.read_ods(filename, has_header=False, sheet_id=0)
            else:
                QErrorMessage(self).showMessage("Invalid file type")
                return
        except Exception as e:
            QErrorMessage(self).showMessage(str(e))
            return

        self.window().stack.insertWidget(1, SpreadsheetInfoWidget(df, filename))
        self.window().stack.setCurrentIndex(1)


class PolarsTableModel(QAbstractTableModel):
    def __init__(self, df: pl.DataFrame, name: str):
        super().__init__()
        self._df = df
        self.name = name

    def rowCount(self, parent=QModelIndex()):
        return self._df.height

    def columnCount(self, parent=QModelIndex()):
        return self._df.width

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        if role == Qt.ItemDataRole.DisplayRole:
            return str(self._df.item(index.row(), index.column()))

        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return self._df.columns[section]
            else:
                return str(section + 1)

        return None


class SpreadsheetInfoWidget(QWidget):
    def __init__(self, df: pl.DataFrame | dict[str, pl.DataFrame], filename: str):
        super().__init__()

        layout = QVBoxLayout(self)

        self.tables = []
        tabs = QTabWidget()
        layout.addWidget(tabs)

        if not isinstance(df, dict):
            df = {filename: df}

        for sheet_name, sheet_df in df.items():
            table = QTableView()
            model = PolarsTableModel(sheet_df, sheet_name)
            table.setModel(model)
            table.selectionModel().selectionChanged.connect(self._on_selection_changed)

            header = table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
            header.setStretchLastSection(True)
            table.resizeColumnsToContents()

            self.tables.append(table)
            tabs.addTab(table, sheet_name)

        self.label = QLabel("Number of selected URLs: 0")
        layout.addWidget(self.label)

        self.buttons_layout = QHBoxLayout()
        layout.addLayout(self.buttons_layout)
        self.back_button = QPushButton("Back")
        self.buttons_layout.addWidget(self.back_button)
        self.back_button.clicked.connect(self._on_back_clicked)
        self.next_button = QPushButton("Next")
        self.buttons_layout.addWidget(self.next_button)
        self.next_button.clicked.connect(self._on_next_clicked)

    def _on_back_clicked(self):
        self.window().stack.setCurrentIndex(0)

    def _on_next_clicked(self):
        selected_items = []
        selected_tab_names = []

        for table in self.tables:
            model = table.model()
            if not (indexes := table.selectionModel().selectedIndexes()):
                continue
            selected_tab_names.append(model.name)
            for index in indexes:
                selected_items.append(model.data(index))

        if not selected_items:
            QErrorMessage(self).showMessage("No items selected")
            return

        # For future widget
        self.all_tabs = len(selected_tab_names) == len(self.tables)
        self.selected_tabs = selected_tab_names

        self.window().stack.insertWidget(2, ListURLWidget(selected_items))
        self.window().stack.setCurrentIndex(2)

    def _on_selection_changed(self):
        selected_items = []

        for table in self.tables:
            model = table.model()
            for index in table.selectionModel().selectedIndexes():
                selected_items.append(model.data(index))

        self.label.setText(f"Number of selected URLs: {len(selected_items)}")


class ListURLWidget(QWidget):
    def __init__(self, urls: list[str]):
        super().__init__()

        layout = QVBoxLayout(self)

        self.label = QLabel(f"Number of URLs: {len(urls)}")
        layout.addWidget(self.label)

        self.edit_buttons_layout = QHBoxLayout()
        layout.addLayout(self.edit_buttons_layout)

        self.add_button = QPushButton("Add")
        self.edit_buttons_layout.addWidget(self.add_button)
        self.add_button.clicked.connect(self._on_add_clicked)

        self.edit_button = QPushButton("Edit")
        self.edit_buttons_layout.addWidget(self.edit_button)
        self.edit_button.clicked.connect(self._on_edit_clicked)

        self.remove_button = QPushButton("Remove")
        self.edit_buttons_layout.addWidget(self.remove_button)
        self.remove_button.clicked.connect(self._on_remove_clicked)

        self.list_widget = QListWidget()
        self.list_widget.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.list_widget.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )

        self.list_widget.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )

        self.list_widget.setAlternatingRowColors(True)
        self.list_widget.setWordWrap(True)

        self.list_widget.model().rowsInserted.connect(self._update_label)
        self.list_widget.model().rowsRemoved.connect(self._update_label)
        self.list_widget.model().dataChanged.connect(self._update_label)
        self.list_widget.model().rowsMoved.connect(self._update_label)

        self.list_widget.addItems(urls)
        for item in (self.list_widget.item(i) for i in range(self.list_widget.count())):
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)

        layout.addWidget(self.list_widget)

        self._delete_shortcut = QShortcut(
            QKeySequence(QKeySequence.StandardKey.Delete), self.list_widget
        )
        self._delete_shortcut.activated.connect(self._on_remove_clicked)

        self.buttons_layout = QHBoxLayout()
        layout.addLayout(self.buttons_layout)
        self.back_button = QPushButton("Back")
        self.buttons_layout.addWidget(self.back_button)
        self.back_button.clicked.connect(self._on_back_clicked)
        self.next_button = QPushButton("Next")
        self.buttons_layout.addWidget(self.next_button)
        self.next_button.clicked.connect(self._on_next_clicked)

        self._update_label()

    def items(self) -> list[str]:
        return [
            self.list_widget.item(i).text() for i in range(self.list_widget.count())
        ]

    def _update_label(self, *args):
        self.label.setText(f"Number of URLs: {self.list_widget.count()}")

    def _current_item_or_warn(self):
        item = self.list_widget.currentItem()
        if item is None:
            QMessageBox.information(self, "Edit item", "Select an item first.")
            return None
        return item

    def _on_add_clicked(self):
        text, ok = QInputDialog.getText(self, "Add URL", "URL:")
        if not ok:
            return
        text = (text or "").strip()
        if not text:
            return
        self.list_widget.addItem(text)
        self._update_label()

    def _on_edit_clicked(self):
        item = self._current_item_or_warn()
        if item is None:
            return

        text, ok = QInputDialog.getText(
            self,
            "Edit URL",
            "URL:",
            text=item.text(),
        )
        if not ok:
            return
        text = (text or "").strip()
        if not text:
            return
        item.setText(text)
        self._update_label()

    def _on_remove_clicked(self):
        selected = self.list_widget.selectedItems()
        if not selected:
            QMessageBox.information(self, "Remove items", "Select one or more items.")
            return

        reply = QMessageBox.question(
            self,
            "Remove items",
            f"Remove {len(selected)} item(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        for item in selected:
            row = self.list_widget.row(item)
            self.list_widget.takeItem(row)

        self._update_label()

    def _on_back_clicked(self):
        self.window().stack.setCurrentIndex(1)

    def _on_next_clicked(self):
        items = self.items()

        if not items:
            QErrorMessage(self).showMessage("No items in the list")
            return

        self.window().stack.insertWidget(
            3, GrabbingImagesWidget(self.window().stack.widget(0).filename, items)
        )
        self.window().stack.setCurrentIndex(3)


@enum
class FailureReason:
    SSL_Error = Case(exc=BaseException)
    InvalidURL = Case(exc=BaseException)
    TimeoutOnNavigation = Case(exc=BaseException)
    NoResponseOrUnreachable = Case(exc=BaseException)
    FailedToLoad = Case(exc=BaseException)
    ResourceNotFound = Case(exc=BaseException)
    InternalServerError = Case(exc=BaseException)
    Unknown = Case(exc=BaseException)
    BlankPage = Case()
    Manual = Case()


class RunParams(TypedDict):
    save_dir: Path
    user_data_dir: Path


class SignalHolder(QObject):
    update_signal = Signal()
    remove_signal = Signal()
    ensure_visible = Signal()
    remove_in_parent_signal = Signal(int)
    add_to_retry_signal = Signal(int)


@dataclasses.dataclass
class ScreenshotObjectData:
    url: str
    inner: Result[Path, FailureReason] = dataclasses.field()


@dataclasses.dataclass
class ScreenshotObject(QObject):
    url: str
    inner: Result[Path, FailureReason] = dataclasses.field()
    retry: bool = dataclasses.field(default=False)
    to_retry_item_added: bool = dataclasses.field(init=False)
    list_index: int | None = dataclasses.field(default=None)
    signal_holder: SignalHolder = dataclasses.field(default_factory=SignalHolder)

    def to_data(self) -> ScreenshotObjectData:
        return ScreenshotObjectData(self.url, self.inner)

    def moveToThread(self, thread, /):
        self.signal_holder.moveToThread(thread)
        super().moveToThread(thread)

    @property
    def ensure_visible(self):
        return self.signal_holder.ensure_visible

    @property
    def update_signal(self):
        return self.signal_holder.update_signal

    @property
    def remove_signal(self):
        return self.signal_holder.remove_signal

    @property
    def add_to_retry_signal(self):
        return self.signal_holder.add_to_retry_signal

    @property
    def remove_in_parent_signal(self):
        return self.signal_holder.remove_in_parent_signal

    def __post_init__(self):
        self.to_retry_item_added = self.retry
        super().__init__()

    @Slot()
    def remove_object(self):
        self.remove_signal.emit()
        self.remove_in_parent_signal.emit(self.list_index)
        del self

    @Slot()
    def toggle_retry(self):
        self.retry = not self.retry
        if not self.to_retry_item_added and self.retry:
            self.add_to_retry_signal.emit(self.list_index)
            self.to_retry_item_added = True
        self.update_signal.emit()


class Runner(QObject):
    add_signal = Signal()
    result_ready = Signal(list)

    def __init__(self, user_data_dir: Path, save_dir: Path):
        super().__init__()

        self.user_data_dir = user_data_dir
        self.save_dir = save_dir

    @Slot(list, QThread)
    def grab_images(self, urls: list[str], receiver_thread: QThread) -> None:
        ret_list: list[ScreenshotObject] = []
        if os.environ.get("PLAYWRIGHT_FIREFOX_POLICIES_JSON") is None:
            os.environ["PLAYWRIGHT_FIREFOX_POLICIES_JSON"] = str(
                LOCAL_DATA_DIR / Path("firefox") / Path("policies.json")
            )
        with psa.sync_playwright() as p:
            browser = p.firefox.launch_persistent_context(
                user_data_dir=self.user_data_dir,
                headless=False,
                viewport={"width": 1920, "height": 1080},
                slow_mo=300,
                firefox_user_prefs={
                    "media.volume_scale": "0.0",
                }
            )
            for url in urls:
                ret_list.append(
                    ScreenshotObject(
                        url,
                        inner := self.get_image(url, browser),
                        retry=inner.is_err(),
                    )
                )
        for obj in ret_list:
            obj.moveToThread(receiver_thread)
        # Move ourselves back to the receiver (main) thread while we're
        # still on the worker thread that owns us, since the main thread
        # can't move us itself
        self.moveToThread(receiver_thread)
        self.result_ready.emit(ret_list)

    def get_image(
        self,
        url: str,
        browser: psa.Browser | psa.BrowserContext,
    ) -> Result[Path, FailureReason]:
        page = browser.new_page()
        response = self.page_goto(page, url)
        if response.is_err():
            page.close()
            self.add_signal.emit()
            return response  # ty:ignore[invalid-return-type]

        self.thread().msleep(1000)

        if "instagram.com" in url:
            page.get_by_role("button", name="close").click()
            self.thread().msleep(500)

        if "youtube.com" in url or "youtu.be" in url:
            self.thread().msleep(4000)

        if "linkedin.com/posts" in url:
            page.get_by_role("button", name="dismiss").click()
            self.thread().msleep(500)

        if "facebook.com" in url and ("groups" in url or "videos" in url or "reel" in url):
            page.get_by_role("button", name="close").click()
            self.thread().msleep(500)

        if "tiktok.com" in url:
            self.thread().msleep(2000)

        page.screenshot(
            path=(pth := self.save_dir / Path(f"{time.time_ns()}.jpeg")),
            type="jpeg",
            quality=70,
        )
        page.close()
        self.add_signal.emit()
        return Result.Ok(pth)

    def page_goto(
        self, page: psa.Page, url: str
    ) -> Result[psa.Response, FailureReason]:
        try:
            response = page.goto(url, timeout=60000)
        except psa.Error as e:
            if "net::ERR_SSL_PROTOCOL_ERROR" in e.message:
                return Result.Err(FailureReason.SSL_Error(exc=e))
            elif "net::ERR_NAME_NOT_RESOLVED" in e.message:
                return Result.Err(FailureReason.InvalidURL(exc=e))
            elif "Navigation timeout of" in e.message:
                return Result.Err(FailureReason.TimeoutOnNavigation(exc=e))
            elif (
                "net::ERR_CONNECTION_REFUSED" in e.message
                or "net::ERR_CONNECTION_TIMED_OUT" in e.message
            ):
                return Result.Err(FailureReason.NoResponseOrUnreachable(exc=e))
            elif "net::ERR_FAILED" in e.message:
                return Result.Err(FailureReason.FailedToLoad(exc=e))
            elif "net::ERR_FILE_NOT_FOUND" in e.message:
                return Result.Err(FailureReason.ResourceNotFound(exc=e))
            elif "net::ERR_ABORTED" in e.message:
                return Result.Err(FailureReason.InternalServerError(exc=e))
            else:
                return Result.Err(FailureReason.Unknown(exc=e))
        if response is None:
            return Result.Err(FailureReason.BlankPage())
        if response.status == 404:
            return Result.Err(FailureReason.ResourceNotFound(exc=None))
        if response.status == 500:
            return Result.Err(FailureReason.InternalServerError(exc=None))
        return Result.Ok(response)


class GrabbingImagesWidget(QWidget):
    operate_worker = Signal(list, QThread)
    worker_thread = QThread()

    @overload
    def __init__(self, filename: str, urls: list[str]): ...
    @overload
    def __init__(self, runner: Runner, obj_list: list[ScreenshotObject]): ...
    def __init__(
        self, filename: str | Runner, urls: list[str] | list[ScreenshotObject]
    ):
        super().__init__()

        self.base_list: list[ScreenshotObject] | None = None

        if isinstance(filename, str):
            work_urls: list[str] = urls

            save_dir = (
                LOCAL_DIR
                / Path("runs")
                / Path(
                    f"{os.path.basename(filename)}_{datetime.datetime.now():%Y-%m-%d_%H-%M-%S}"
                )
            )
            save_dir.mkdir(parents=True, exist_ok=False)
            user_data_dir = (
                LOCAL_DIR / Path("browser_data") / Path("firefox") / Path("default")
            )
            atexit.register(
                shutil.rmtree,
                save_dir,
                ignore_errors=True,
                onerror=None,
                onexc=None,
                dir_fd=None,
            )

            self.runner = Runner(user_data_dir, save_dir)

        else:
            self.base_list: list[ScreenshotObject] = urls
            work_urls: list[str] = [obj.url for obj in self.base_list if obj.retry]

            self.runner: Runner = filename

        self.runner.moveToThread(self.worker_thread)
        self.operate_worker.connect(self.runner.grab_images)
        self.runner.result_ready.connect(self.result_ready)
        self.runner.add_signal.connect(self.add_progress)

        layout = QVBoxLayout(self)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, len(work_urls))
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.current_val = 0

        self.buttons_layout = QHBoxLayout()
        layout.addLayout(self.buttons_layout)
        self.back_button = QPushButton("Cancel")
        self.buttons_layout.addWidget(self.back_button)
        self.back_button.clicked.connect(self._on_back_clicked)

        self.worker_thread.start()
        self.operate_worker.emit(work_urls, self.thread())

    def closeEvent(self, event, /):
        self.worker_thread.quit()
        self.worker_thread.wait()
        super().closeEvent(event)

        # If we were called from ListURLWidget, our Runner will not be
        # reused and should be deleted
        if getattr(self, "runner", None) is not None and self.base_list is None:
            self.runner.deleteLater()

    @Slot()
    def add_progress(self):
        self.current_val += 1
        self.progress_bar.setValue(self.current_val)

    @Slot(list)
    def result_ready(self, result: list[ScreenshotObject]):
        if self.base_list is None:
            self.result = result
        else:
            self.result = self.base_list
            for i, obj in enumerate(self.base_list):
                if not obj.retry:
                    continue
                self.result[i].inner = (rob := result.pop(0)).inner
                self.result[i].retry = rob.retry

        self.runner.add_signal.disconnect(self.add_progress)
        self.runner.result_ready.disconnect(self.result_ready)
        self._on_next_clicked()

    def _on_back_clicked(self):
        self.window().stack.setCurrentIndex(2)
        self.window().stack.removeWidget(self)
        self.deleteLater()

    def _on_next_clicked(self):
        # If called from ListURLWidget, we are at 3 =>
        # we replace this widget; users shouldn't be able to go back
        # to this widget
        # If called for retries by ResultPreviewWidget, we are at 4 =>
        # We delete the old widget first
        if self.window().stack.currentIndex() == 4:
            self.window().stack.removeWidget(widg := self.window().stack.widget(3))
            widg.deleteLater()
        self.window().stack.insertWidget(
            3, ResultPreviewWidget(self.result, self.runner)
        )
        self.window().stack.setCurrentIndex(3)
        self.deleteLater()


class OverlayIconEngine(QIconEngine):
    def __init__(self, base: QIcon, overlay: QIcon, base_alpha: float = 1):
        super().__init__()
        self._base = base
        self._overlay = overlay
        self._base_alpha = base_alpha

    def clone(self):
        return OverlayIconEngine(self._base, self._overlay, self._base_alpha)

    def paint(
        self, painter: QPainter, rect: QRect, mode: QIcon.Mode, state: QIcon.State
    ):
        painter.save()
        painter.setOpacity(self._base_alpha)
        self._base.paint(painter, rect, Qt.AlignmentFlag.AlignCenter, mode, state)
        painter.restore()

        self._overlay.paint(painter, rect, Qt.AlignmentFlag.AlignCenter, mode, state)

    def pixmap(self, size, mode, state):
        # Render a real composed pixmap (super().pixmap() doesn't know how to draw our layers).
        size = QSize(size)
        if not size.isValid():
            return QPixmap()

        # Prefer the screen DPR when available (prevents HiDPI scaling artifacts).
        app = QApplication.instance()
        screen = app.primaryScreen() if app else None
        dpr = float(screen.devicePixelRatio() if screen else 1.0)

        pm = QPixmap(int(size.width() * dpr), int(size.height() * dpr))
        pm.setDevicePixelRatio(dpr)
        pm.fill(Qt.GlobalColor.transparent)

        p = QPainter(pm)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)

        rect = QRect(0, 0, size.width(), size.height())
        self.paint(p, rect, mode, state)
        p.end()

        return pm


class _PixmapFitLabel(QWidget):
    """A widget that paints a pixmap scaled to fit (keeps aspect ratio).

    Uses paintEvent instead of QLabel.setPixmap to avoid the
    setPixmap → sizeHint change → layout → resizeEvent → setPixmap
    infinite-recursion cycle.
    """

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._src: QPixmap | None = None
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)

    def setSourcePixmap(self, pixmap: QPixmap):
        self._src = pixmap
        self.update()  # schedule a repaint – no layout side-effects

    def paintEvent(self, event):
        if self._src is None or self._src.isNull():
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        scaled = self._src.scaled(
            self.size(),
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        x = (self.width() - scaled.width()) // 2
        y = (self.height() - scaled.height()) // 2
        painter.drawPixmap(x, y, scaled)
        painter.end()


class _AspectRatioWidget(QWidget):
    """A container widget that enforces a fixed aspect ratio (width / height).

    Uses resizeEvent + setFixedHeight so the layout always allocates the
    correct height, even inside scroll-areas or layouts that ignore
    heightForWidth().
    """

    def __init__(self, aspect_w: int, aspect_h: int, parent: QWidget | None = None):
        super().__init__(parent)
        self._aspect_w = aspect_w
        self._aspect_h = aspect_h
        self._resizing = False
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._resizing:
            return
        target_h = int(event.size().width() * self._aspect_h / self._aspect_w)
        if self.height() != target_h:
            self._resizing = True
            self.setFixedHeight(target_h)
            self._resizing = False


class _SquareButton(QPushButton):
    """A QPushButton that stays square (1:1 aspect ratio)."""

    def __init__(self, *args, max_side: int = 48, **kwargs):
        super().__init__(*args, **kwargs)
        self._max_side = max_side
        self._resizing = False
        self.setMaximumSize(max_side, max_side)
        sp = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        sp.setHeightForWidth(True)
        self.setSizePolicy(sp)

    def hasHeightForWidth(self) -> bool:
        return not self._resizing

    def heightForWidth(self, w: int) -> int:
        if self._resizing:
            return -1
        return min(w, self._max_side)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._resizing:
            return
        side = min(event.size().width(), event.size().height(), self._max_side)
        if self.width() != side or self.height() != side:
            self._resizing = True
            self.setFixedSize(side, side)
            self._resizing = False


class ScreenshotPreviewWidget(QWidget):
    def __init__(self, screenshot_object: ScreenshotObject):
        super().__init__()

        self.retry_icon = qta.icon("fa6s.rotate")
        self.retry_checked_icon = QIcon(
            OverlayIconEngine(
                qta.icon(
                    "fa6s.rotate",
                    color="grey",
                ),
                qta.icon(
                    "fa6s.check",
                    color="green",
                ),
            )
        )

        self.screenshot_object = screenshot_object
        self.screenshot_object.update_signal.connect(self.update_widget)
        self.screenshot_object.remove_signal.connect(self.remove_self)
        self.screenshot_object.ensure_visible.connect(self.ensure_visible)

        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Minimum)

        pre_layout = QVBoxLayout(self)
        pre_layout.setContentsMargins(0, 0, 0, 0)
        pre_layout.setSpacing(4)

        # --- URL row (just overflows / scrolls horizontally) ---
        self.url_label = QLabel()
        self.url_label.setWordWrap(False)
        self.url_label.setTextFormat(Qt.TextFormat.RichText)
        self.url_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextBrowserInteraction
        )
        self.url_label.setOpenExternalLinks(True)
        self.url_label.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed
        )
        self.url_label.setFixedHeight(self.url_label.fontMetrics().height())

        self.url_scroll = QScrollArea()
        self.url_scroll.setWidget(self.url_label)
        self.url_scroll.setWidgetResizable(True)
        self.url_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.url_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.url_scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.url_scroll.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        self.url_scroll.setFixedHeight(self.url_label.fontMetrics().height() + 2)


        # --- Image + buttons row (19:9 aspect ratio) ---
        self._aspect_container = _AspectRatioWidget(19, 9)
        layout = QHBoxLayout(self._aspect_container)
        layout.setContentsMargins(0, 0, 0, 0)
        pre_layout.addWidget(self._aspect_container)
        pre_layout.addWidget(self.url_scroll)

        self.image_container = QVBoxLayout()
        self.image_container.setContentsMargins(0, 0, 0, 0)
        layout.addLayout(self.image_container, stretch=1)

        self.displayed = screenshot_object.inner
        self.display_image()

        self.replace_button = _SquareButton(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton), ""
        )
        self.replace_button.clicked.connect(self.replace_image)
        self.remove_button = _SquareButton(
            self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon), ""
        )
        self.remove_button.clicked.connect(self.screenshot_object.remove_object)
        self.retry_state = self.screenshot_object.retry
        self.retry_button = _SquareButton(
            self.retry_icon if not self.retry_state else self.retry_checked_icon, ""
        )
        self.retry_button.clicked.connect(self.screenshot_object.toggle_retry)

        self.buttons_layout = QVBoxLayout()
        self.buttons_layout.setContentsMargins(0, 0, 0, 0)
        self.buttons_layout.setSpacing(8)
        self.buttons_layout.addWidget(self.replace_button)
        self.buttons_layout.addWidget(self.remove_button)
        self.buttons_layout.addWidget(self.retry_button)
        layout.addLayout(self.buttons_layout)

        self._set_url(self.screenshot_object.url)

    @Slot()
    def ensure_visible(self):
        self.parent().parent().parent().ensureWidgetVisible(self)

    def _set_url(self, url: str):
        safe_url = (url or "").strip()
        self.url_label.setText(f'<a href="{safe_url}">{safe_url}</a>')

    @Slot()
    def replace_image(self):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptOpen)
        dialog.setNameFilter(
            "Supported Images ("
            + " ".join(
                map(
                    lambda barr: "*." + barr.data().decode(),
                    QImageReader.supportedImageFormats(),
                )
            )
            + ")"
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            path = Path(dialog.selectedFiles()[0])
            self.screenshot_object.inner = Result.Ok(path)
            self.screenshot_object.update_signal.emit()

    def display_image(self):
        while (widget := self.image_container.takeAt(0)) is not None:
            self.image_container.removeWidget(widget.widget())
            widget.widget().deleteLater()
        if self.displayed.is_ok():
            pixmap = QPixmap(self.displayed.unwrap())
            image = _PixmapFitLabel(self)
            image.setSourcePixmap(pixmap)
            self.image_container.addWidget(image)
        else:
            reason = self.displayed.unwrap_err()
            failure_text_map: dict[type[FailureReason], str] = {
                FailureReason.SSL_Error: "SSL Error",
                FailureReason.InvalidURL: "Invalid URL",
                FailureReason.TimeoutOnNavigation: "Timeout on navigation",
                FailureReason.NoResponseOrUnreachable: "No response or unreachable",
                FailureReason.FailedToLoad: "Failed to load",
                FailureReason.ResourceNotFound: "Resource not found",
                FailureReason.InternalServerError: "Internal Server Error",
                FailureReason.Unknown: "Unknown error",
                FailureReason.BlankPage: "Blank page",
                FailureReason.Manual: "Manual Removal",
            } # ty:ignore[invalid-assignment]
            cross = QWidget()
            cross.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            cross_layout = QHBoxLayout(cross)
            cross_layout.setContentsMargins(0, 0, 0, 0)
            cross_layout.setSpacing(4)
            cross_icon = QLabel(pixmap=self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxCritical).pixmap(32, 32))
            cross_text = QLabel(failure_text_map[type(reason)])
            cross_layout.addWidget(cross_icon)
            cross_layout.addWidget(cross_text)
            self.image_container.addWidget(cross)

    @Slot()
    def update_widget(self):
        if self.screenshot_object.inner != self.displayed:
            self.displayed = self.screenshot_object.inner
            self.display_image()
        if self.screenshot_object.retry != self.retry_state:
            self.retry_state = self.screenshot_object.retry
            self.retry_button.setIcon(
                self.retry_icon if not self.retry_state else self.retry_checked_icon
            )
        if self.screenshot_object.url != self.url_label.text():
            self._set_url(self.screenshot_object.url)

    @Slot()
    def remove_self(self):
        self.layout().removeWidget(self)
        self.deleteLater()


class ScreenshotObjListWidgetItem(QListWidgetItem):
    def __init__(self, obj: ScreenshotObject):
        super().__init__(obj.url)

        self.obj = obj
        self.obj.update_signal.connect(self.update_self)
        self.obj.remove_signal.connect(self.remove_self)

        self.setFlags(
            self.flags()
            | Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsUserCheckable
        )
        self.setCheckState(
            Qt.CheckState.Checked if self.obj.retry else Qt.CheckState.Unchecked
        )

    def __lt__(self, other: ScreenshotObjListWidgetItem) -> bool:
        return self.obj.list_index < other.obj.list_index

    @Slot()
    def checkbox_triggered(self):
        if self.obj.retry == (self.checkState() == Qt.CheckState.Checked):
            # update_self triggered itemChanged
            return
        self.obj.toggle_retry()

    @Slot()
    def update_self(self):
        if self.obj.retry != (self.checkState() == Qt.CheckState.Checked):
            self.setCheckState(
                Qt.CheckState.Checked if self.obj.retry else Qt.CheckState.Unchecked
            )
        if self.obj.url != self.text():
            self.setText(self.obj.url)

    def changed(self):
        if self.obj.url != self.text():
            self.obj.url = self.text()
            self.obj.update_signal.emit()
        if self.obj.retry != (self.checkState() == Qt.CheckState.Checked):
            self.obj.toggle_retry()

    @Slot()
    def remove_self(self):
        lw = self.listWidget()
        if lw:
            lw.takeItem(lw.row(self))
            del self


class ResultPreviewWidget(QWidget):
    def __init__(self, results: list[ScreenshotObject], runner_for_retry: Runner):
        super().__init__()

        self.runner_for_retry = runner_for_retry

        self.results = results
        for i, obj in enumerate(self.results):
            obj.list_index = i
            obj.add_to_retry_signal.connect(self.add_retry)
            obj.remove_in_parent_signal.connect(self.remove_obj)

        self.main_layout = QHBoxLayout(self)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.scroll_area.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.main_layout.addWidget(self.scroll_area)

        self._images_container = QWidget()
        self.images_layout = QVBoxLayout(self._images_container)
        self._images_container.setLayout(self.images_layout)
        self.scroll_area.setWidget(self._images_container)

        for obj in self.results:
            self.images_layout.addWidget(ScreenshotPreviewWidget(obj))
        self.images_layout.addStretch(1)

        self.retry_layout = QVBoxLayout()
        self.main_layout.addLayout(self.retry_layout)
        label = QLabel("Retry List")
        self.retry_layout.addWidget(label)

        self.list_widget = QListWidget()
        self.list_widget.setSortingEnabled(True)
        self.retry_layout.addWidget(self.list_widget)
        self.list_widget.setDragDropMode(QAbstractItemView.DragDropMode.NoDragDrop)
        self.list_widget.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.list_widget.setSizePolicy(
            QSizePolicy.Policy.MinimumExpanding, QSizePolicy.Policy.Expanding
        )

        self.list_widget.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.list_widget.itemClicked.connect(self.item_clicked)

        self.list_widget.setWordWrap(True)
        self.list_widget.model().dataChanged.connect(self.data_change)

        # to_retry_item_added is True if object is to be retried
        # also True if it was just retried
        for obj in filter(lambda ob: ob.to_retry_item_added, results):
            item = ScreenshotObjListWidgetItem(obj)
            self.list_widget.addItem(item)

        self.buttons_layout = QHBoxLayout()
        self.back_button = QPushButton("Back to URL list")
        self.back_button.clicked.connect(self._on_back_selected)
        self.buttons_layout.addWidget(self.back_button)
        self.retry_button = QPushButton("Retry Selected")
        self.retry_button.clicked.connect(self._on_retry_selected)
        self.buttons_layout.addWidget(self.retry_button)
        self.next_button = QPushButton("Create PDF")
        self.next_button.clicked.connect(self._on_next_selected)
        self.buttons_layout.addWidget(self.next_button)

        self.retry_layout.addLayout(self.buttons_layout)

    @Slot()
    def item_clicked(self, item: ScreenshotObjListWidgetItem):
        item.obj.ensure_visible.emit()

    def _on_back_selected(self):
        self.window().stack.setCurrentIndex(2)
        self.window().stack.removeWidget(self)
        self.deleteLater()

    def _on_retry_selected(self):
        if sum(obj.retry for obj in self.results) == 0:
            QMessageBox.information(self, "No items to retry", "No items to retry")
            return
        self.window().stack.insertWidget(
            4, GrabbingImagesWidget(self.runner_for_retry, self.results)
        )
        self.window().stack.setCurrentIndex(4)

    def _on_next_selected(self):
        self.window().stack.insertWidget(
            4, WaitOnPDFCreationWidget([obj.to_data() for obj in self.results])
        )
        self.window().stack.setCurrentIndex(4)

    def remove_obj(self, obj_index: int):
        self.images_layout.removeWidget(self.images_layout.itemAt(obj_index).widget())
        del self.results[obj_index]

    def add_retry(self, obj_index: int):
        item = ScreenshotObjListWidgetItem(self.results[obj_index])
        self.list_widget.addItem(item)

    @Slot(QModelIndex, QModelIndex, list)
    def data_change(
        self, topLeft: QModelIndex, bottomRight: QModelIndex, roles: list[int]
    ):
        for i in range(topLeft.row(), bottomRight.row() + 1):
            item: None | QListWidgetItem = self.list_widget.item(i)
            if item is None:
                continue
            item.changed()

    def remove_object(self, i: int):
        del self.results[i]


class PDFRunner(QObject):
    add_signal = Signal()
    result_ready = Signal(bytes)

    def create_pdf(self, data: list[ScreenshotObjectData]):
        pdf_stream = io.BytesIO()
        c = canvas.Canvas(pdf_stream, pagesize=landscape(A4), pageCompression=1)

        c.setTitle("")
        c.setAuthor("")
        c.setSubject("")
        c.setCreator("ScraPDF")

        width, height = landscape(A4)

        for obj in data:
            if obj.inner.is_err():
                self.add_signal.emit()
                continue

            with Image.open(obj.inner.unwrap()) as im:
                im = im.convert("RGB")
                im = im.resize(
                    (1280, 720), resample=Image.Resampling.LANCZOS, reducing_gap=2
                )

                c.drawImage(
                    ImageReader(im),
                    0,
                    0,
                    width=width,
                    height=height,
                    preserveAspectRatio=True,
                )

            wid = c.stringWidth(obj.url)
            c.drawString(10, 10, obj.url, mode=0)
            c.linkURL(
                obj.url, rect=(10, 10 - 2, min(10 + wid, width), 10 + c._fontsize)
            )
            c.showPage()

            self.add_signal.emit()

        # Leaving the people to wait at 100%
        c.save()
        self.result_ready.emit(pdf_stream.getvalue())


class WaitOnPDFCreationWidget(QWidget):
    operate_worker = Signal(list)
    worker_thread = QThread()

    def __init__(self, data: list[ScreenshotObjectData]):
        super().__init__()

        self.runner = PDFRunner()

        self.runner.moveToThread(self.worker_thread)
        self.operate_worker.connect(self.runner.create_pdf)
        self.runner.result_ready.connect(self.result_ready)
        self.runner.add_signal.connect(self.add_progress)

        layout = QVBoxLayout(self)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, len(data))
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.current_val = 0

        self.buttons_layout = QHBoxLayout()
        layout.addLayout(self.buttons_layout)
        self.back_button = QPushButton("Cancel")
        self.buttons_layout.addWidget(self.back_button)
        self.back_button.clicked.connect(self._on_back_clicked)

        self.worker_thread.start()
        self.operate_worker.emit(data)

    def closeEvent(self, event, /):
        self.worker_thread.quit()
        self.worker_thread.wait()
        super().closeEvent(event)

        if getattr(self, "runner", None) is not None:
            self.runner.deleteLater()

    @Slot()
    def add_progress(self):
        self.current_val += 1
        self.progress_bar.setValue(self.current_val)

    @Slot(bytes)
    def result_ready(self, result: bytes):
        self.result = result
        self._on_next_clicked()

    def _on_back_clicked(self):
        self.window().stack.setCurrentIndex(3)
        self.window().stack.removeWidget(self)
        self.deleteLater()

    def _on_next_clicked(self):
        self.window().stack.insertWidget(4, FinalScreen(self.result))
        self.window().stack.setCurrentIndex(4)
        self.deleteLater()


class FinalScreen(QWidget):
    def __init__(self, pdf_bytes: bytes):
        super().__init__()

        self.pdf_bytes = pdf_bytes

        self.optimize = Qt.CheckState.Checked

        layout = QVBoxLayout(self)

        self.checkbox = QCheckBox("Optimize PDF")
        layout.addWidget(self.checkbox, alignment=Qt.AlignmentFlag.AlignCenter)
        self.checkbox.setCheckState(self.optimize)
        self.checkbox.checkStateChanged.connect(self.trigger_optimize)

        self.button = QPushButton("Save PDF")
        layout.addWidget(self.button, alignment=Qt.AlignmentFlag.AlignCenter)
        self.button.clicked.connect(self._on_save_clicked)

        self.exit = QPushButton("Exit")
        layout.addWidget(self.exit, alignment=Qt.AlignmentFlag.AlignCenter)
        self.exit.clicked.connect(QApplication.instance().quit)

    @Slot(bool)
    def trigger_optimize(self, state):
        self.optimize = state

    def _on_save_clicked(self):
        original = self.window().stack.widget(0).path_fname
        all_tabs = self.window().stack.widget(1).all_tabs
        selected_tabs = self.window().stack.widget(1).selected_tabs

        pdf_name = original.with_suffix("").name + "_".join([""] + selected_tabs if not all_tabs else []) + ".pdf"

        dialog = QFileDialog(self)
        dialog.setDefaultSuffix("pdf")
        dialog.selectFile(pdf_name)
        dialog.setDirectory(str(original.parent))
        dialog.setNameFilter("PDF (*.pdf)")
        dialog.setFileMode(QFileDialog.FileMode.AnyFile)
        dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)

        if not dialog.exec() == QDialog.DialogCode.Accepted:
            return

        path = Path(dialog.selectedFiles()[0])

        written = path.write_bytes(self.pdf_bytes)

        if written != len(self.pdf_bytes):
            QErrorMessage(self).showMessage(
                f"Something failed whilst writing to {path}\nFile might be corrupted!"
            )
            return

        if self.optimize == Qt.CheckState.Checked:
            pdf = pike.open(path, allow_overwriting_input=True)
            if len(pdf.pages) > 0:
                # If someone feels funny enough to create an empty PDF
                pdf.save(path, linearize=True)
            pdf.close()

        success = QMessageBox(self)
        success.setIcon(QMessageBox.Icon.Information)
        success.setWindowTitle("Success")
        success.setTextFormat(Qt.TextFormat.RichText)
        success.setText(f'PDF saved to\n<a href="{path.as_uri()}">{path}</a>')
        success.setStandardButtons(QMessageBox.StandardButton.Ok)
        success.exec()


# https://stackoverflow.com/a/60358637
class DiagnosticStyle(QProxyStyle):
    def polish(self, widget):
        super().polish(widget)
        if isinstance(widget, QWidget):
            widget.setAttribute(Qt.WidgetAttribute.WA_Hover, True)

    def drawControl(self, element, option, painter, widget=None):
        # First let the base style draw normally
        super().drawControl(element, option, painter, widget)

        if widget is None or painter is None:
            return

        # Border around the widget
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        try:
            if QApplication.widgetAt(QCursor.pos()) is widget:
                pen = QPen(QColor("red"))
                pen.setWidth(4)
                painter.setPen(pen)
                painter.drawRect(widget.rect())
                # Translucent overlay
                translucent = QBrush(QColor(255, 246, 240, 100))
                painter.fillRect(widget.rect(), translucent)

                # Class name text
                class_name = widget.metaObject().className()
                fm = painter.fontMetrics()
                text_rect = fm.boundingRect(
                    widget.rect(),
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                    class_name,
                )
                bg_rect = text_rect.adjusted(-2, -1, 2, 1)
                painter.fillRect(bg_rect, QColor(255, 255, 255, 220))
                painter.setPen(QPen(QColor("darkblue")))
                painter.drawText(
                    widget.rect(),
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                    class_name,
                )
            elif widget.underMouse():
                pen = QPen(QColor("red"))
                pen.setWidth(2)
                painter.setPen(pen)
                painter.drawRect(widget.rect())
            else:
                pen = QPen(QColor("red"))
                pen.setWidth(1)
                painter.setPen(pen)
                painter.drawRect(widget.rect())
        finally:
            painter.restore()


def main():
    app = QApplication(sys.argv)
    if __debug__:
        app.setStyle(DiagnosticStyle())

    main_window = MainWindow()
    main_window.show()

    app.exec()


if __name__ == "__main__":
    main()
