import googlesearch
import requests
import bs4
import pandas as pd


class GoogleSearch:
    def __init__(self):
        self.search = input('Enter a search term: ')

    def get_links(self):
        '''Gets the links of top Google search results, and appends them to a list'''
        search = googlesearch.search(self.search, tld='com', num=10, stop=1, pause=2)
        results = []
        for link in search:
            results.append(link)
        return results

    def get_results(self, results):
        '''Saves search related information to a dictionary'''
        page_info = {
                    'title': [],
                    'query': [],
                    'position': [],
                    'link': [],
                    }
        position = 1
        for link in results:
            print('Retrieving info for result {}...'.format(position))
            site = requests.get(link)
            content = bs4.BeautifulSoup(site.text, 'lxml')
            page_title = content.title.string
            try:
                page_info['title'].append(page_title)
            except Exception as e:
                page_info['title'].append('<<NO TITLE FOUND>>')
            page_info['query'].append(self.search)
            page_info['position'].append(position)
            page_info['link'].append(link)
            position += 1
        return page_info

    def save_to_file(self, page_info):
        '''Converts page_info dict into a dataframe and saves to an excel file'''
        df = pd.DataFrame.from_dict(page_info)
        df.to_excel('search-results_{}.xlsx'.format(self.search).replace(' ', '-'))
        print('Results have been saved to a file in your working directory')
        return df


def main():
    my_search = GoogleSearch()
    top_links = my_search.get_links()
    search_results = my_search.get_results(top_links)
    my_search.save_to_file(search_results)


if __name__ == '__main__':
    main()
