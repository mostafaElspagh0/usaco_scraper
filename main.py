from selenium import webdriver
from dataclasses import dataclass
import xlwt
from xlwt import Workbook


@dataclass
class Problem:
    name: str
    judge: str
    difficulty: str
    link: str
    star: bool
    topic: str

    def __key(self):
        return self.name, self.judge, self.difficulty, self.link, self.star, self.topic

    def __hash__(self):
        return hash(self.__key())

    def __eq__(self, other):
        if isinstance(other, Problem):
            return self.__key() == other.__key()
        return NotImplemented


def get_all_unique_problems(page_url: str, driver: webdriver.chrome):
    problems = []
    driver.get(page_url)
    topic = driver.find_element_by_tag_name('h1').text
    for i in driver.find_elements_by_tag_name('tr'):
        if i.get_attribute('id').startswith('problem'):
            tds = i.find_elements_by_tag_name('td')
            problems.append(Problem(tds[2].text.strip(), tds[1].text.strip(),
                                 tds[3].text.strip(), tds[2].find_element_by_tag_name('a').get_attribute('href'),
                                 len(tds[2].find_elements_by_tag_name('svg')) != 0, topic))

    for i in driver.find_elements_by_class_name("border-t-4"):
        ps = i.find_elements_by_tag_name("p")
        problems.append(Problem(ps[0].text, ps[1].text.split('-')[0].strip(), ps[1].text.split('-')[1].strip(),
                             i.find_element_by_tag_name('a').get_attribute('href'), False, topic))
    return problems


def write_problem_at_row(p: Problem, she: xlwt.Worksheet, curr: int):
    style = xlwt.easyxf('font: bold 1, color red;')
    if p.star:
        she.write(curr, 0, xlwt.Formula(f'HYPERLINK("{p.link}";"{p.name}")'), style)
    else:
        she.write(curr, 0, xlwt.Formula(f'HYPERLINK("{p.link}";"{p.name}")'))
    she.write(curr, 1, p.judge)
    she.write(curr, 2, p.topic)
    she.write(curr, 4, p.difficulty)


def get_all_links(main_link: str, driver: webdriver.chrome):
    driver.get(main_link)
    driver.implicitly_wait(1)
    return [anchor.get_attribute('href')
            for anchor in
            driver.find_element_by_xpath("//div[@class='flex-1" +
                                         " h-0 overflow-y-auto']").find_elements_by_tag_name('a')]


chrome_options = webdriver.ChromeOptions()
prefs = {"profile.managed_default_content_settings.images": 2}
chrome_options.add_experimental_option("prefs", prefs)
d = webdriver.Chrome(options=chrome_options)
wb = Workbook()
px = []

generalSheet = wb.add_sheet('general', cell_overwrite_ok=True)
bronzeSheet = wb.add_sheet('bronze', cell_overwrite_ok=True)
silverSheet = wb.add_sheet('silver', cell_overwrite_ok=True)
goldSheet = wb.add_sheet('gold', cell_overwrite_ok=True)
platinumSheet = wb.add_sheet('platinum', cell_overwrite_ok=True)
advancedSheet = wb.add_sheet('advanced', cell_overwrite_ok=True)


links = [
    ("https://usaco.guide/general/using-this-guide", generalSheet),
    ("https://usaco.guide/bronze/time-comp", bronzeSheet),
    ('https://usaco.guide/silver/binary-search-sorted', silverSheet),
    ('https://usaco.guide/gold/divis', goldSheet),
    ('https://usaco.guide/plat/seg-ext', platinumSheet),
    ('https://usaco.guide/adv/springboards', advancedSheet)
]

tcount = 70

for link, sheet in links:
    tcount += 1
    for i in get_all_links(link, d):
        px = px + get_all_unique_problems(i, d)
    cu = 1
    for i in px:
        write_problem_at_row(i, sheet, cu)
        cu += 1
    px.clear()
    wb.save(f'{tcount}.xls')

d.close()
wb.save('final.xls')
