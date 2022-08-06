import matplotlib.pyplot as plt
import wordcloud as wc


class wordFig:
    def __init__(self, text='', exclude=[]):
        self.stopwords = wc.STOPWORDS
        self.text = str(text)
        self.loaded = False
        for word in exclude:
            self.stopwords.add(word)
        if str(self.text) != '':
            self.generate()

    def generate(self):
        self.fig = wc.WordCloud(background_color='white',
                                width=1500,
                                height=960,
                                margin=10,
                                stopwords=self.stopwords
                                ).generate(self.text)
        self.loaded = True

    def addText(self, text):
        self.text += text

    def addExclude(self, exclude):
        for word in exclude:
            self.stopwords.add(word)

    def showPng(self, title='WordCloud'):
        plt.figure().canvas.set_window_title(title)
        # plt.title(title)
        plt.axis('off')
        try:
            plt.imshow(self.fig)
        except Exception as e:
            print(e)
            print('图片加载失败！请尝试重新生成后加载。')
        plt.show()

    def savePng(self, name):
        self.fig.to_file(name)
