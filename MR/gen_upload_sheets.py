import pandas as pd
import openpyxl


def rename_sheets(file_path):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(file_path)

    # Iterate through all sheets in the workbook
    for sheet_name in workbook.sheetnames:
        #        if ':' in sheet_name:
        # print(sheet_name)
        try:
            # Extract the part of the sheet name after ":"
            # new_sheet_name = sheet_name.split(':')[1].strip()
            new_sheet_name = sheet_name.strip().split(' ')[1].strip()
            # print(new_sheet_name)

            # Rename the sheet
            workbook[sheet_name].title = new_sheet_name
        except:
            pass

    # Save the modified workbook
    workbook.save(file_path)


class Analyze_question:
    def __init__(self, wfh_alldata_path, curation_path) -> None:
        self.alldata = pd.read_csv(wfh_alldata_path)
        self.curation = pd.ExcelFile(curation_path)
        upload_sheet_name = wfh_alldata_path.split('-')[0].strip() + '.xlsx'
        # self.result=pd.ExcelWriter('WFH_Upload.xlsx',engine='xlsxwriter')
        self.result = pd.ExcelWriter(upload_sheet_name, engine='xlsxwriter')

        self.alldata.to_excel(
            self.result, sheet_name='all-topics', index=False)

        for question in self.curation.sheet_names[1:]:
            response_df = self.curation.parse(
                question, usecols=[0, 1, 2], skiprows=[0])
            topic_sub_df = self.curation.parse(
                question, usecols=[4, 5], skiprows=[0]).dropna()

            insights_df = self.curation.parse(
                question, usecols=[8, 10, 11], skiprows=[0]).dropna()
            rep_df = self.rep_texts(topic_sub_df)

            ex_rep_df = rep_df.join(insights_df.rename(
                columns={'Topic.2': 'title'}).set_index('title'), on='title', how='left')

            self.rep_texts_export(ex_rep_df, question)
            rep_df.set_index('title', inplace=True)
            topics = rep_df[rep_df['P_id'] == -1000]
            subtopics = rep_df.reset_index().drop_duplicates(
                subset='title', keep='last').set_index('title')
            # print(topics)
            res_df = response_df.join(topics['labels'], on='Topic', how='left')
            res_df = res_df.rename(
                columns={'labels': question+'_label_0', 'Topic': question+'_title_0'})

            res_df = res_df.join(
                subtopics['labels'], on='Subtopic', how='left')
            
            print("#####",res_df.columns, res_df.columns[0])
            # res_df = res_df.rename(columns={
            #                        'labels': question+'_label_1', 'Response': question, 'Subtopic': question+'_title_1'})
            
            res_df = res_df.rename(columns={
                                   'labels': question+'_label_1', res_df.columns[0]: question, 'Subtopic': question+'_title_1'})

            print("#####Updated",res_df.columns)
            res_df[question+'_proximity_score_0'] = 0.5
            res_df[question+'_proximity_score_1'] = 0.5
            
            print("###Columns in res_df:", res_df.columns)


            res_df = res_df[[question, question + '_label_0', question + '_title_0', question + '_proximity_score_0', question + '_label_1', question + '_title_1', question + '_proximity_score_1']]
            res_df.set_index(question, inplace=True)

            self.alldata = self.alldata.join(res_df, on=question, how='left')
            

            self.alldata = self.alldata.reset_index().drop_duplicates(
                subset='index', keep='first').set_index('index')

            columns_to_replace = [question +'_label_0', question +'_title_0', question +
                                  '_proximity_score_0', question + '_label_1', question + '_title_1', question + '_proximity_score_1']
            replacement_value = {
                question+'_label_0': topics.loc['Low Content',]['labels'],
                question+'_title_0': 'Low Content',
                question+'_proximity_score_0': 0.5,
                question+'_label_1': topics.loc['Low Content',]['labels'],
                question+'_title_1': 'Low Content',
                question+'_proximity_score_1': 0.5,
            }
            self.alldata.loc[1:len(self.alldata[question].dropna())-1, columns_to_replace] = self.alldata.loc[1:len(
                self.alldata[question].dropna())-1, columns_to_replace].fillna(replacement_value)

        self.alldata.to_excel(
            self.result, sheet_name='all-topics', index=False)

        self.result.close()

    def __str__(self) -> str:
        return f'Upload sheet created successful'

    def rep_texts(self, topic_sub_df):
        neg_id = topic_sub_df.loc[topic_sub_df['Topic.1'].isin(
            ['Low Content'])]['Topic.1']
        neg_id = pd.concat([neg_id, pd.Series(['Outliers'])])

        res_df = pd.DataFrame()
        pos_id = topic_sub_df[~ topic_sub_df['Topic.1'].isin(
            ['Low Content'])]['Topic.1'].drop_duplicates()
        res_df['title'] = pd.concat([neg_id, pos_id]).reset_index(drop=True)
        res_df['P_id'] = -1000
        res_df['labels'] = [k for k in range(-1*len(neg_id), len(pos_id))]
        for topic in pos_id:
            label = res_df['labels'].tail(1).values[0]
            pid = res_df[res_df['title'] == topic]['labels'].values[0]
            rdf = topic_sub_df.loc[topic_sub_df['Topic.1']
                                   == topic].reset_index(drop=True)
            rdf['P_id'] = pid
            rdf['labels'] = [i+label+1 for i in rdf.index.values]
            rdf = rdf.drop(columns=['Topic.1'])
            rdf = rdf.rename(columns={"Subtopic.1": "title"})
            res_df = pd.concat([res_df, rdf]).reset_index(drop=True)
        return res_df

    def rep_texts_export(self, res_df, question):
        #print("Columns in res_df:", res_df.columns, res_df.columns[0])
        
        res_df = res_df.rename(columns={'labels': 'label',
                                        'P_id': 'parent',
                                        res_df.columns[3]: 'exemplar_statement_0',
                                        res_df.columns[4]: 'exemplar_statement_1'})
        # res_df = res_df.rename(columns={'labels': 'label',
        #                                 res_df.columns[1]: 'parent',
        #                                 res_df.columns[3]: 'exemplar_statement_0',
        #                                 res_df.columns[4]: 'exemplar_statement_1'})
        #print("Columns in res_df:", res_df.columns)
        res_df['proximity_score'] = 0.5
        res_df['summary'] = None
        res_df['keywords'] = None
        res_df['count'] = 5
        res_df['percentage'] = 5
        res_df['cohesion_score'] = .5
        res_df['exemplar_statement_0_score'] = .5
        res_df['exemplar_statement_1_score'] = .5

        res_df = res_df[['label', 'parent',
                         'title',
                        'proximity_score',
                         'summary',
                         'keywords',
                         'count',
                         'percentage',
                         'cohesion_score',
                         'exemplar_statement_0',
                         'exemplar_statement_0_score',
                         'exemplar_statement_1',
                         'exemplar_statement_1_score']]
        res_df.fillna('None', inplace=True)
        res_df.to_excel(self.result, sheet_name=question +
                        '-rep-texts', index=False)


if __name__ == "__main__":
    # Provide the path to your Excel file
    curation_sheet = 'Market Research-Curation.xlsx'
    all_data = 'Market Research Upload - all-topics.csv'
    # rename sheets
    rename_sheets(curation_sheet)

    obj = Analyze_question(all_data, curation_sheet)
    print(obj)
