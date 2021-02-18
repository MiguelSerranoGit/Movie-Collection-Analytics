import json
from os import listdir
import pandas as pd
import requests
from analitycs import charts, convert_csv_into_excel, highlight_excel

csv_collection = 'collection.csv'
excel_file = 'collection.xlsx'
arrow = ' --> '


def home():
    print('\nWELCOME TO THE MAIN MENU!\nWhat do you want to do?')
    print('\t1) Search for a movie in the database\n'
          '\t2) Search for a movie in my collection\n'
          '\t3) Exit\n'
          'Enter an option:')
    menu(int(input()))


def menu(option):
    while option != str(1) or option != str(2) or option != str(3):
        if option == 1:
            print('''You have chosen the option 'Search for a movie in the database'  ''')
            search_movie()
        elif option == 2:
            print('''You have chosen the option 'Search for a movie in my collection' ''')
            search_movie_in_my_collection()
        elif option == 3:
            print('''You have chosen the option 'Exit'
                \nWe are creating the excel, making it pretty and making the graphs. Then everything will be ready ''')
            convert_csv_into_excel()
            charts()
            highlight_excel()
            print('¡See you soon!')
            exit(0)
        print('The option you have chosen is not valid. Do it again')
        home()

def search_movie_in_my_collection():
    data = pd.read_csv(csv_collection)
    titulos = data['TITLE']
    print('Write the title of the movie:')
    titulo_buscar = input()
    resultados = []
    for element in range(0, len(titulos)):
        if titulos[element].lower().find(titulo_buscar.lower()) != -1:
            resultados.append(element)
    if len(resultados) > 0:
        print('We have found ' + str(len(resultados)) + ' results')
        for element in resultados:
            print(data.iloc[element])
            print('------------------------------')
        home()
    else:
        print('No results found')
        search_movie()


def create_csv():
    df = pd.DataFrame(columns=('ID', 'TITLE', 'ORIGINAL TITLE', 'SCORE', 'YEAR', 'BUDGET', 'GROSS', 'BENEFIT',
                               'CATEGORY', 'RUNTIME'))
    df.to_csv(csv_collection, index=False, encoding='utf-8-sig')
    print('csv file successfully created.')


def check_csv(option):
    if csv_collection not in listdir('.'):
        print('The csv file does not exist, it must be created.')
        create_csv()
        if option == 1:
            print('''You don't have any movies saved yet. Add some first. ''')
            home()
        elif option == 2:
            print('Now its time to add the film')
    else:
        collection = pd.read_csv(csv_collection)
        if not collection.empty:
            if option == 1:
                home()
        else:
            if option == 1:
                print('''You don't have any titles saved yet. Add a movie before''')
                home()
            elif option == 2:
                print('The file exists, but it is empty. Now you have to add the movie')


def search_movie():
    print('Type the title of the movie you want to search for')
    url_first_part = 'https://api.themoviedb.org/3/search/movie?api_key=YOUR_API_KEY&language=es' \
                     '&query= '
    query = input()
    if query.__contains__(' '):
        query = query.replace(' ', '%20')
    url_final = '''&page=1&include_adult=false'''
    url_completed = url_first_part + query + url_final
    movie_request = requests.get(url_completed)
    search_results = json.loads(movie_request.content)
    number_of_results = search_results['total_results']
    if number_of_results > 1:
        print('There are many results, it will be necessary to make a more detailed search. \n'
              'These are all the results we have found:')
        for element in search_results['results']:
            if 'release_date' not in element:
                release_date = 'Coming soon'
                print(element['title'] + arrow + str(element['id']) + arrow
                      + release_date + arrow + element['overview'])
            else:
                if element['release_date'] == '':
                    element['release_date'] = 'Coming soon'
                print(element['title'] + arrow + str(element['id']) + arrow
                      + str(element['release_date']) + arrow + element['overview'])
        print('\nHave you found the movie you are looking for?\n1) YES\n2) NO')
        answer_movie_found = input()
        if answer_movie_found == str(1):
            print('\nOk, so copy its ID to use it in the search:')
            id_movie = input()
            url_start1 = 'https://api.themoviedb.org/3/movie/' + str(id_movie)
            url_final1 = '?api_key=YOUR_API_KEY&language=es'
            url_completed1 = url_start1 + url_final1
            movie_request1 = requests.get(url_completed1)
            search_result = json.loads(movie_request1.content)
            show_info(search_result, 2)
        else:
            print('\nSpecify a little more:')
            search_movie()
    elif number_of_results == 0:
        print('No results found. Search again')
        search_movie()
    else:
        for element in search_results['results']:
            print(element['title'] + arrow + str(element['id']) + arrow
                  + str(element['release_date']) + arrow + element['overview'])
        print('\nCopy the ID of the film to use it in your search:')
        id_movie = input()
        url_start1 = 'https://api.themoviedb.org/3/movie/' + str(id_movie)
        url_final1 = '?api_key=YOUR_API_KEY&language=es'
        url_completed1 = url_start1 + url_final1
        movie_request1 = requests.get(url_completed1)
        search_result = json.loads(movie_request1.content)
        show_info(search_result, 2)


def show_info(movie, option):
    if option == 1:
        movie_info = movie['results']
        print('This is the information of the movie:')
        movie_title = movie_info[0]['title']
        print('\tTitle: %s' % movie_title)
        movie_original_title = movie_info[0]['original_title']
        print('\tOriginal title: %s' % movie_original_title)
        identifier = movie_info[0]['id']
        print('\tID: %s' % str(identifier))
        budget_movie = movie_info[0]['budget']
        print('\tBudget: %s' % budget_movie)
        gross_movie = movie_info[0]['revenue']
        print('\tGross: %s' % gross_movie)
        benefit_movie = gross_movie - budget_movie
        print('\tProfit: %s' % benefit_movie)
        release_year = movie_info[0]['release_date'][0:4]
        print('\tYear: %s' % release_year)
        score = movie_info[0]['vote_average']
        print('\tScore: %s' % str(score))
        runtime = movie['runtime']
        print('\tDuration: %s' % str(runtime))
    if option == 2:
        identifier = movie['id']
        print('\tID: %s' % str(identifier))
        movie_title = movie['title']
        print('\tTitle: %s' % movie_title)
        movie_original_title = movie['original_title']
        print('\tOriginal title: %s' % movie_original_title)
        budget_movie = movie['budget']
        print('\tBudget: %s' % str(budget_movie))
        gross_movie = movie['revenue']
        print('\tGross: %s' % str(gross_movie))
        benefit_movie = gross_movie - budget_movie
        print('\tProfit: %s' % str(benefit_movie))
        release_year = movie['release_date'][0:4]
        print('\tYear: %s' % release_year)
        score = movie['vote_average']
        print('\tScore: %s' % str(score))
        runtime = movie['runtime']
        print('\tDuration: %s' % str(runtime))
    print('Do you want to add it?\n1) YES\n2) NO')
    answer1 = input()
    if int(answer1) == 1:
        print('Okay, lets add it')
        check_csv(2)
        add_movie(identifier, movie_title, movie_original_title, score, budget_movie, gross_movie,
                  benefit_movie, release_year, runtime)
        print('Do you want to do something else?\n1) YES\n2) NO')
        answer3 = input()
        while answer3 != str(1) or answer3 != str(2):
            if int(answer3) == 1:
                convert_csv_into_excel()
                home()
            elif int(answer3) == 2:
                convert_csv_into_excel()
                charts()
                highlight_excel()
                print('See you soon!')
                exit(0)
            print('The option you have written is not valid. Do it again.')
            answer3 = input()
    elif int(answer1) == 2:
        print('The movie has not been added. Exit to the main menu')
        home()


def add_movie(identifier, movie_title, movie_original_title, score, budget, gross, benefit, release_year, runtime):
    global category
    check_movie_in_list(identifier)
    print('Have you seen the movie, or do you have the movie pending?\n1) VIEWED\n2) PENDING')
    respuesta2 = input()
    if int(respuesta2) == 1:
        category = 'Vista'
    elif int(respuesta2) == 2:
        category = 'Pendiente'
    movie_info = {'ID': [int(identifier)], 'TITLE': [str(movie_title)], 'ORIGINAL TITLE': [str(movie_original_title)],
                  'YEAR': [int(release_year)], 'SCORE': [float(score)], 'BUDGET': [budget],
                  'GROSS': [gross], 'BENEFIT': [benefit], 'CATEGORY': [category], 'RUNTIME': [runtime]}
    df2 = pd.DataFrame(movie_info)
    csv_file = pd.read_csv(csv_collection)
    edited_dataframe = csv_file.append(df2, sort=False)
    edited_dataframe.to_csv(csv_collection, index=False, encoding='utf-8-sig')
    print('File successfully edited\n')


def check_movie_in_list(identifier):
    print('First of all lets see if it is already in your collection.')
    csv = pd.read_csv(csv_collection)
    for element in csv['ID']:
        if element == identifier:
            print('The movie you are looking for has already been saved.')
            home()
    print('The film is not saved')


if __name__ == '__main__':
    home()
