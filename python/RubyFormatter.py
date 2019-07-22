import uno
import unohelper
import gettext
import os
import re
import urllib.request
from urllib.parse import urlparse
from com.sun.star.beans import PropertyValue
from com.sun.star.datatransfer import XTransferable
from com.sun.star.datatransfer import DataFlavor
from com.sun.star.datatransfer import UnsupportedFlavorException
from com.sun.star.task import XJobExecutor

# To match Han characters
han_regex = re.compile("[\u2E80-\u2FDF々〇〻\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF\u20000-\u2FFFF]+")
# To match Kana characters
kana_regex = re.compile("[\u3041-\u308F\u30A1-\u30FA・ー･ｰ]+")
# To match Side points
points_regex = re.compile("[,.、··•∙⋅・･\s]+")
point_regex = re.compile("[,.、··•∙⋅・･]")
# To replace Double angle bracket
dab_regex = re.compile("《")

IMP_NAME = "io.github.hallelujahdrive.rubyformatter"

KAKUYOMU = 0
PIXIV = 1
NAROU = 2

class TextList(list):
    def __init__(self, format):
        self.format = format


    # Replace double angle bracket and append to list
    def append_word(self, word):
        if self.format == KAKUYOMU or self.format == NAROU:

            if len(self) > 0 and self.__is_filled_han(self[-1]):
                word = dab_regex.sub("|《", word)

        self.append(word)

    # Attention:
    # If a ruby text contains "《"　or "》", Kakuyomu and Shousetsuka ni narou viewer do not format the text correctly.
    def append_word_and_ruby(self, word, ruby_text):
        if self.format == KAKUYOMU:
            if self.__is_side_points(word, ruby_text):
                self.append("《《")
                self.append(word)
                self.append("》》")
            else:
                self.__append_word(word)
                self.__append_ruby(ruby_text)

        elif self.format == PIXIV:
            self.append("[[rb:")
            self.append(word)
            self.append(" > ")
            self.append(ruby_text)
            self.append("]]")

        elif self.format == NAROU:
            if self.__is_side_points(word, ruby_text):
                for c in word:
                    self.__append_word(c)
                    self.__append_ruby("・")
            
            else:
                self.__append_word(word)
                self.__append_ruby(ruby_text)


    def __append_word(self, word):
        if not self.__is_filled_han(word):
            self.append("|")
        self.append(word)


    def __append_ruby(self, ruby_text):
        self.append("《")
        self.append(ruby_text)
        self.append("》")        


    def __is_filled_han(self, word):
        return han_regex.fullmatch(word) is not None

    
    def __is_filled_kana(self, word):
        return lana_regex.fullmatch(word) is not None


    def __is_side_points(self, word, ruby_text):
        if points_regex.fullmatch(ruby_text) is None:
            return False

        res = point_regex.findall(ruby_text)
        return res is not None and len(res) == len(word)


class RubyFormatter(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def __create_dialog(self):
        try:
            translater = gettext.translation(
                "messages",
                localedir = urllib.request.url2pathname(
                    get_package_location(self.ctx, IMP_NAME)
                    + "locales"),
                languages = [get_language(self.ctx)],
                fallback = False
            )
        except Exception as e:
            translater = gettext.translation(
                "messages",
                localedir = urllib.request.url2pathname(
                    get_package_location(self.ctx, IMP_NAME)
                    + "locales"),
                languages = ["en"],
                fallback=True
            )

        translater.install()
        _ = translater.gettext

        format_list = [
            _("Kakuyomu"),
            _("Pixiv"),
            _("Shosetsuka ni naro")
            ]

        dp = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", self.ctx)
        dlg = dp.createDialog("vnd.sun.star.extension://io.github.hallelujahdrive.rubyformatter/dialogs/RubyFormatterDialog.xdl")
        dlg_model = dlg.Model

        dlg_model.Title =  _("Select ruby format and copy")

        format_list_box = dlg_model.getByName("FormatListBox")
        format_list_box.StringItemList = format_list
        format_list_box.SelectedItems = [0]

        OK_button = dlg_model.getByName("OKButton")
        OK_button.Label = _("OK")
        
        cancel_button = dlg_model.getByName("CancelButton")
        cancel_button.Label = _("Cancel")

        return dlg


    def trigger(self, args):
        dlg = self.__create_dialog()
        if not dlg.execute() == 1:
            return

        list_box = dlg.Model.getByName("FormatListBox")
        selected = list_box.SelectedItems[0]
        
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        doc = desktop.getCurrentComponent()

        dlg.dispose()
        
        copy_to_clipboard(format_text(doc, selected))


def get_language(ctx):
    cp = ctx.ServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
    prop = PropertyValue()
    prop.Name = "nodepath"
    prop.Value = "org.openoffice.Setup/L10N"
    key = "ooLocale"
    try:
        ca = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (prop,))

        if ca and ca.hasByName(key):
            lang = ca.getPropertyValue(key)
            return lang
    except:
        pass

    return ""


def get_package_location(ctx, module_name):
    pip = ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
    return urlparse(pip.getPackageLocation(module_name)).path + "/"


def format_text(doc, selected):
    text = TextList(selected)

    lines_enum = doc.getText().createEnumeration()
    while lines_enum.hasMoreElements():
        line_elem = lines_enum.nextElement()
        words_enum = line_elem.createEnumeration()
        if len(text) > 0:
            text.append("\n")
        
        ruby_text = None
        while words_enum.hasMoreElements():
            word_elem = words_enum.nextElement()
            if ruby_text is None:
                text.append_word(word_elem.String)
            else:
                text.append_word_and_ruby(word_elem.String, ruby_text)

            ruby_text = word_elem.RubyText

    return "".join(text)


def copy_to_clipboard(text):
    ctx = uno.getComponentContext()
    sc = ctx.ServiceManager.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)
    sc.setContents(TextTransferable(text), None)


class TextTransferable(unohelper.Base, XTransferable):
    def __init__(self, text):
        self.text = text
        self.unicode_content_type = "text/plain;charset=utf-16"

    
    def getTransferData(self, flavor):
        if flavor.MimeType.lower() != self.unicode_content_type:
            raise UnsupportedFlavorException()
        
        return self.text
    

    def getTransferDataFlavors(self):
        return DataFlavor(MimeType=self.unicode_content_type, HumanPresentableName="Unicode Text"),


    def isDataFlavorSuppoerted(self, flavor):
        return flavor.MimeType.lower() == self.unicode_content_type


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    RubyFormatter,
    IMP_NAME,
    ("com.sun.star.task.Job",),)