import pandas as pd
class Analyze_question:
    def __init__(self,wfh_alldata_path,curation_path) -> None:
        self.alldata=pd.read_csv(wfh_alldata_path)
        self.curation=pd.ExcelFile(curation_path)
        self.result=pd.ExcelWriter('WFH_Upload.xlsx',engine='xlsxwriter')
        self.alldata.to_excel(self.result,sheet_name='all-topics',index=False)
        
        for question in self.curation.sheet_names[1:]:
            response_df=self.curation.parse(question,usecols=[0,1,2],skiprows=[0])
            topic_sub_df=self.curation.parse(question,usecols=[4,5],skiprows=[0]).dropna()
            
            insights_df= self.curation.parse(question,usecols=[8,10,11],skiprows=[0]).dropna()
            rep_df=self.rep_texts(topic_sub_df)

            ex_rep_df= rep_df.join(insights_df.rename(columns={'Topic.2':'title'}).set_index('title'),on='title',how='left')


            self.rep_texts_export(ex_rep_df,question)
            rep_df.set_index('title',inplace=True)
            topics=rep_df[rep_df['P_id']==-1000]
            subtopics=rep_df.reset_index().drop_duplicates(subset='title', keep='last').set_index('title')
            
            res_df=response_df.join(topics['labels'],on='Topic',how='left')
            res_df=res_df.rename(columns={'labels':question+'_label_0','Topic':question+'_title_0'})
            
            res_df=res_df.join(subtopics['labels'],on='Subtopic',how='left')
            res_df=res_df.rename(columns={'labels':question+'_label_1','Response':question,'Subtopic':question+'title_1'})
            
            res_df[question+'_proximity_score_0']=0.5
            res_df[question+'_proximity_score_1']=0.5
            res_df=res_df[[question,question+'_label_0', question+'_title_0',question+'_proximity_score_0',question+'_label_1', question+'title_1',question+'_proximity_score_1']]
            res_df.set_index(question,inplace=True)

            self.alldata=self.alldata.join(res_df,on=question,how='left')
            self.alldata=self.alldata.reset_index().drop_duplicates(subset='index', keep='first').set_index('index')

            
        self.alldata.to_excel(self.result,sheet_name='all-topics',index=False)

        self.result.close()
            
        
    def __str__(self) -> str:
        return f'exported successful'
    
    def rep_texts(self,topic_sub_df):
        neg_id=topic_sub_df.loc[topic_sub_df['Topic.1'].isin(['Low Content'])]['Topic.1']
        neg_id=pd.concat([neg_id,pd.Series(['Outliers'])])

        res_df= pd.DataFrame()
        pos_id=topic_sub_df[~ topic_sub_df['Topic.1'].isin(['Low Content'])]['Topic.1'].drop_duplicates()
        res_df['title']=pd.concat([neg_id,pos_id]).reset_index(drop=True)
        res_df['P_id']=-1000
        res_df['labels']=[ k for k in range(-1*len(neg_id),len(pos_id))]
        for  topic in pos_id:
            label=res_df['labels'].tail(1).values[0]
            pid=res_df[res_df['title']==topic]['labels'].values[0]
            rdf= topic_sub_df.loc[topic_sub_df['Topic.1']==topic].reset_index(drop=True)
            rdf['P_id']=pid
            rdf['labels']=[i+label+1 for i in rdf.index.values]
            rdf=rdf.drop(columns=['Topic.1'])
            rdf=rdf.rename(columns={"Subtopic.1":"title"})
            res_df=pd.concat([res_df,rdf]).reset_index(drop=True)
        return res_df
    
    def rep_texts_export(self,res_df,question):
        res_df=res_df.rename(columns={'labels':'label',
                       'P_id':'parent',
                       'Exemplar_1':'exemplar_statement_0',
                       'Exemplar_2':'exemplar_statement_1'})
        res_df['proximity_score']=0.5
        res_df['summary']=None
        res_df['keywords']=None
        res_df['count']=5
        res_df['percentage']=5
        res_df['cohesion_score']=.5
        res_df['exemplar_statement_0_score']=.5
        res_df['exemplar_statement_1_score']=.5

        res_df=res_df[['label','parent',
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
        res_df.fillna('None',inplace=True)
        res_df.to_excel(self.result,sheet_name= question+'-rep-texts' ,index=False)


obj=Analyze_question('WFH upload - all-topics.csv', 'WFH Curation.xlsx')
print(obj)