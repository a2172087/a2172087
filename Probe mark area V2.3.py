import sys
import math
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel, 
                             QFileDialog, QGraphicsScene, QGraphicsView, 
                             QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsItem, QMessageBox, QDesktopWidget)
from PyQt5.QtGui import QPixmap, QPen, QIcon, QFont 
from PyQt5.QtCore import Qt, QRectF, QPointF, QSizeF
import qtmodern.styles
import qtmodern.windows

class ResizableRectItem(QGraphicsRectItem):
    def __init__(self, rect):
        super().__init__(rect)
        self.setFlags(QGraphicsItem.ItemIsMovable | QGraphicsItem.ItemIsSelectable)
        self.dragging = False
        self.resizing = False

    def mousePressEvent(self, event):
        if self.edgeContains(event.pos()):
            self.dragging = True
        elif event.button() == Qt.RightButton:
            self.resizing = True
            self.startPos = event.pos()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.dragging:
            newRect = self.rect()
            newRect.setBottomRight(event.pos())
            self.setRect(newRect)
        elif self.resizing:
            newRect = QRectF(self.rect().topLeft(), event.pos()).normalized()
            self.setRect(newRect)
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.dragging = False
        if self.resizing and event.button() == Qt.RightButton:
            self.resizing = False
        super().mouseReleaseEvent(event)

    def edgeContains(self, point):
        edgeMargin = 10.0
        rect = self.rect()
        if QRectF(rect.right() - edgeMargin, rect.bottom() - edgeMargin, edgeMargin, edgeMargin).contains(point):
            return True
        return False

class ResizableEllipseItem(QGraphicsEllipseItem):
    def __init__(self, rect):
        super().__init__(rect)
        self.setFlags(QGraphicsItem.ItemIsMovable | QGraphicsItem.ItemIsSelectable)
        self.resizing = False
        self.rotating = False
        self.startAngle = 0

    def mousePressEvent(self, event):
        if event.button() == Qt.RightButton:
            self.resizing = True
            self.startPos = event.pos()
        elif event.modifiers() & Qt.ShiftModifier:
            self.rotating = True
            self.startAngle = self.calculateAngle(event.pos())
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.resizing:
            newRect = QRectF(self.rect().topLeft(), event.pos()).normalized()
            self.setRect(newRect)
        elif self.rotating:
            angle = self.calculateAngle(event.pos())
            rotation = angle - self.startAngle
            self.setTransformOriginPoint(self.boundingRect().center())
            self.setRotation(self.rotation() + rotation)
            self.startAngle = angle
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.resizing and event.button() == Qt.RightButton:
            self.resizing = False
        elif self.rotating:
            self.rotating = False
        super().mouseReleaseEvent(event)

    def calculateAngle(self, point):
        center = self.boundingRect().center()
        return math.atan2(point.y() - center.y(), point.x() - center.x()) * 180 / math.pi


class ImageClassifier(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.rectLineThickness = 1  # Default line thickness for rectangles
        self.ellipseLineThickness = 1  # Default line thickness for ellipses

    def initUI(self):
        font = QFont("微軟正黑體", 9)
        QApplication.setFont(font)

        self.setGeometry(500, 500, 1360, 1080)
        self.setWindowTitle('Probe mark area')
        self.setWindowIcon(QIcon('icon_result.ico'))

        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene, self)
        self.view.setGeometry(150, 20, 1200, 1015)

        self.btnSelectImage = QPushButton('選擇圖片', self)
        self.btnSelectImage.setGeometry(10, 20, 130, 30)
        self.btnSelectImage.clicked.connect(self.selectImage)

        self.btnDrawRect = QPushButton('畫方框\n*按右鍵調尺寸', self)
        self.btnDrawRect.setGeometry(10, 75, 130, 80)
        self.btnDrawRect.clicked.connect(self.drawRect)

        self.btnDrawCircle = QPushButton('畫圓圈\n*按Shift可旋轉\n*按右鍵調尺寸', self)
        self.btnDrawCircle.setGeometry(10, 180, 130, 80) 
        self.btnDrawCircle.clicked.connect(self.drawCircle)

        self.btnDelete = QPushButton('刪除', self)
        self.btnDelete.setGeometry(10, 280, 130, 30)
        self.btnDelete.clicked.connect(self.deleteShape)

        self.btnCalculate = QPushButton('計算面積', self)
        self.btnCalculate.setGeometry(10, 320, 130, 30)
        self.btnCalculate.clicked.connect(self.calculateArea)

        self.labelResult = QLabel('面積佔比: ', self)
        self.labelResult.setGeometry(10, 360, 180, 30)

        self.labelResultNextLine = QLabel('', self)
        self.labelResultNextLine.setGeometry(10, 390, 180, 30)

        self.shapes = []
        self.currentPixmap = None

        # 窗口居中設置
        self.centerWindow()

    def centerWindow(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def selectImage(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'Open file', '/home', "Image files (*.jpg *.png *.jpeg)")
        if fname:
            self.currentPixmap = QPixmap(fname)
            self.scene.clear()
            self.scene.addPixmap(self.currentPixmap)
            self.shapes.clear()

    def drawRect(self):
        rectItem = ResizableRectItem(QRectF(100, 100, 200, 200))
        rectItem.setPen(QPen(Qt.red, self.rectLineThickness))  # Use the attribute for line thickness
        self.scene.addItem(rectItem)
        self.shapes.append(rectItem)

    def drawCircle(self):
        ellipseItem = ResizableEllipseItem(QRectF(300, 300, 200, 200))
        ellipseItem.setPen(QPen(Qt.blue, self.ellipseLineThickness))  # Use the attribute for line thickness
        self.scene.addItem(ellipseItem)
        self.shapes.append(ellipseItem)

    def calculateArea(self):
        if not self.shapes:
            self.labelResult.setText("需要至少*1圖形")
            return

        max_area = 0
        min_area = float('inf')
        max_shape = None
        min_shape = None

        for shape in self.shapes:
            area = self.calculateShapeArea(shape)
            if area > max_area:
                max_area = area
                max_shape = shape
            if area < min_area:
                min_area = area
                min_shape = shape

        if not max_shape or not min_shape:
            self.labelResult.setText("無法計算面積")
            return

        ratio = (min_area / max_area) * 100
        self.labelResultNextLine.setText(f'{ratio:.2f}%')

    def calculateShapeArea(self, shape):
        if isinstance(shape, ResizableRectItem):
            return shape.boundingRect().width() * shape.boundingRect().height()
        elif isinstance(shape, ResizableEllipseItem):
            rect = shape.boundingRect()
            return 3.14159 * (rect.width() / 2) * (rect.height() / 2)
        else:
            return 0

    def deleteShape(self):
        selectedItems = self.scene.selectedItems()
        if selectedItems:
            for item in selectedItems:
                self.scene.removeItem(item)
                if item in self.shapes:
                    self.shapes.remove(item)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    qtmodern.styles.dark(app)
    ex = ImageClassifier()
    win = qtmodern.windows.ModernWindow(ex)
    win.show()
    sys.exit(app.exec_())