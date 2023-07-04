import asyncio
import csv
from pyppeteer import launch

async def main():
    browser = await launch(headless=True,executablePath ="C:\Program Files\Google\Chrome\Application\chrome.exe")
    page = await browser.newPage()
    await page.goto('https://fermi.utmb.edu/cgi-bin/new/sdapcgi_01_food.cgi')
    
    #################################
    command = 'document.getElementsByClassName("styled-table")[1].children[1].children.length'
    length = await page.evaluate(command, force_expr=True)
    
    with open('AllergyDatas.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Food', 'Allergen', 'Allergen Level'])
        
        for i in range(0,length):
            command = f'document.getElementsByClassName("styled-table")[1].children[1].children[{i}].children[1].children[0].href'
            link = await page.evaluate(command, force_expr=True)
            items = []
            for j in range(0,6):
                command = f'document.getElementsByClassName("styled-table")[1].children[1].children[{i}].children[{j}].innerText'
                item = await page.evaluate(command, force_expr=True)
                items.append(item)
                if (j == 1):
                    page2 = await browser.newPage()
                    await page2.goto('http://www.allergen.org')
                    
                    await page2.waitForFunction('!(document.getElementById("allergenname") == null)');
    
                                    
                    command = f'document.getElementById("allergenname").value = "{item}"'
                    item = await page2.evaluate(command, force_expr=True)
                    
                    
                    command = f'document.getElementById("simplesearch").children[3].click()'
                    item = await page2.evaluate(command, force_expr=True)
                    
                    page2.waitForNavigation()
    
                    try:
                        await page2.waitForFunction('!(document.getElementById("isotable") == null)',timeout=3000);
                        command = f'document.getElementById("isotable").children[1].children[0].children[3].innerText'
                        item = await page2.evaluate(command, force_expr=True)
                    except :
                        pass  
                    
                    
                    # command = f'document.getElementsByClassName("page_title")[0].innerText == "Search Results: 0"'
                    # item2 = await page2.evaluate(command)
                    # if (item2 == "false"):
                    #     await page2.waitForFunction('!(document.getElementById("isotable") == null)');
                    #     command = f'document.getElementById("isotable").children[1].children[0].children[3].innerText'
                    #     item = await page2.evaluate(command, force_expr=True)
                    # else:
                    #     item = "null"
                    
                    
                    items.append(item)
                    
                    await page2.close()
                    
            
            writer.writerow(items)
    
    await browser.close()

asyncio.get_event_loop().run_until_complete(main())
print('ALL DONE ^__^')
