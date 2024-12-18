# PO-number modifier
import pandas as pd 


def modify():
    # reading the csv file 
    df = pd.read_csv("Farnell.csv")
    newNumber = df.loc[0,'PO Number'] + '-1'
    for index in range(0,len(df)):
        df.loc[index, 'PO Number'] = newNumber
    df.to_csv("Farnell.csv", index=False) 
    

def main():
    modify()


if __name__ == '__main__':
    main()