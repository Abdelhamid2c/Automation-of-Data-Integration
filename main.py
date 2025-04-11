import wx
import random
import os

class CirclePanel(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        self.SetName("1")  # Setting the name of the panel to "1"
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.Bind(wx.EVT_PAINT, self.on_paint)
        
        # Couleurs initiales
        self.circle_color = "yellow"
        self.line_color = "red"
        # Ajout d'un grand cercle
        self.large_circle_color = None  # Sera définie par MainFrame
        self.show_large_circle = False
        
    def on_paint(self, event):
        dc = wx.AutoBufferedPaintDC(self)
        dc.Clear()
        
        # Get the size of the panel
        w, h = self.GetSize()
        center_x, center_y = w // 2, h // 2
        radius = min(w, h) // 2 - 10
        
        # Dessiner le grand cercle si demandé
        if self.show_large_circle and self.large_circle_color:
            dc.SetBrush(wx.Brush(self.large_circle_color))
            dc.SetPen(wx.Pen("black", 1))
            large_radius = radius + 20  # Plus grand que le cercle normal
            dc.DrawCircle(center_x, center_y, large_radius)
            
        # Draw the circle with current color (level 0)
        dc.SetBrush(wx.Brush(self.circle_color))
        dc.SetPen(wx.Pen("black", 1))
        dc.DrawCircle(center_x, center_y, radius)
        
        # Draw the diagonal line with current color (level 1)
        dc.SetPen(wx.Pen(self.line_color, 2))
        dc.DrawLine(center_x - radius, center_y - radius, 
                   center_x + radius, center_y + radius)
    
    def change_colors(self):
        # Liste de couleurs possibles
        colors = ["yellow", "blue", "green", "cyan", "pink", "orange", "purple"]
        line_colors = ["red", "black", "blue", "green", "purple"]
        
        # Choisir des couleurs aléatoires différentes des couleurs actuelles
        new_circle_color = self.circle_color
        while new_circle_color == self.circle_color:
            new_circle_color = random.choice(colors)
        
        new_line_color = self.line_color
        while new_line_color == self.line_color:
            new_line_color = random.choice(line_colors)
        
        self.circle_color = new_circle_color
        self.line_color = new_line_color
        
        # Rafraîchir le panneau pour redessiner avec les nouvelles couleurs
        self.Refresh()
        
    def apply_large_circle(self, color):
        """Applique un grand cercle de la couleur spécifiée"""
        self.large_circle_color = color
        self.show_large_circle = True
        self.Refresh()

class ImagePanel(wx.Panel):
    def __init__(self, parent, image_path):
        super().__init__(parent)
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.Bind(wx.EVT_PAINT, self.on_paint)
        
        # Charger l'image si le chemin est valide
        self.bitmap = None
        if os.path.exists(image_path):
            image = wx.Image(image_path, wx.BITMAP_TYPE_ANY)
            self.bitmap = wx.Bitmap(image)
        
    def on_paint(self, event):
        dc = wx.AutoBufferedPaintDC(self)
        dc.Clear()
        
        # Dessiner l'image si elle a été chargée
        if self.bitmap:
            w, h = self.GetSize()
            img_w, img_h = self.bitmap.GetWidth(), self.bitmap.GetHeight()
            
            # Calculer le ratio pour redimensionner tout en conservant les proportions
            ratio = min(w/img_w, h/img_h)
            new_w, new_h = int(img_w*ratio), int(img_h*ratio)
            
            # Calculer la position centrée
            pos_x = (w - new_w) // 2
            pos_y = (h - new_h) // 2
            
            # Redimensionner et dessiner l'image
            resized_img = self.bitmap.ConvertToImage().Scale(new_w, new_h, wx.IMAGE_QUALITY_HIGH)
            dc.DrawBitmap(wx.Bitmap(resized_img), pos_x, pos_y)

class MainFrame(wx.Frame):
    def __init__(self, image_path=None):
        super().__init__(None, title="Cercle avec Ligne Diagonale et Image", size=(500, 700))
        
        # Définir la couleur du grand cercle basée sur l'image
        self.image_circle_color = "blue"  # Couleur par défaut, à adapter selon votre image
        
        # Créer un sizer vertical pour organiser les éléments
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Ajouter le panneau d'image si un chemin est fourni
        if image_path:
            self.image_panel = ImagePanel(self, image_path)
            main_sizer.Add(self.image_panel, 1, wx.EXPAND | wx.ALL, 10)
        
        # Ajouter le panneau de cercle
        self.circle_panel = CirclePanel(self)
        # Configurer la couleur du grand cercle
        self.circle_panel.large_circle_color = self.image_circle_color
        main_sizer.Add(self.circle_panel, 1, wx.EXPAND | wx.ALL, 10)
        
        # Ajouter le bouton pour changer les couleurs
        self.color_button = wx.Button(self, label="Appliquer le grand cercle")
        self.color_button.Bind(wx.EVT_BUTTON, self.on_apply_large_circle)
        main_sizer.Add(self.color_button, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        
        # Appliquer le sizer
        self.SetSizer(main_sizer)
        self.Show()
    
    def on_apply_large_circle(self, event):
        """Applique le grand cercle au cercle actuel"""
        self.circle_panel.apply_large_circle(self.image_circle_color)
        
    def on_change_colors(self, event):
        self.circle_panel.change_colors()

if __name__ == "__main__":
    app = wx.App()
    
    # Chemin de votre image
    image_path = r"C:\Users\user\Desktop\conn1.png"
    
    frame = MainFrame(image_path)
    app.MainLoop()