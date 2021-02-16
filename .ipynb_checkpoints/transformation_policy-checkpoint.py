import pandas as pd

policy_df=pd.read_csv('WFH Curation - impact_quality.csv')

policy_df.columns=policy_df.iloc[0]
policy_df=policy_df.drop([0]).reset_index(drop=True)

topic_sub_df=policy_df.iloc[:,[4,5]].dropna()

insights_df= policy_df.iloc[:,[8,10,11]].dropna()


neg_id=topic_sub_df.loc[topic_sub_df['Topic'].isin(['Low Content'])]['Topic']
neg_id=neg_id.append(pd.Series(['Outliers']),ignore_index=True)

res_df= pd.DataFrame()

pos_id=topic_sub_df[~ topic_sub_df['Topic'].isin(['Low Content'])]['Topic'].drop_duplicates()
#print(type(pos_id))
res_df['title']=pd.concat([neg_id,pos_id]).reset_index(drop=True)
res_df['P_id']=-1000

res_df['labels']=[ k for k in range(-1*len(neg_id),len(pos_id))]

res_df= res_df.join(insights_df.rename(columns={'Topic':'title'}).set_index('title'),on='title',how='left')

#label=res_df['labels'].tail(1).values[0]
#print(label)
for  topic in pos_id:
    label=res_df['labels'].tail(1).values[0]
    pid=res_df[res_df['title']==topic]['labels'].values[0]
    rdf= topic_sub_df.loc[topic_sub_df['Topic']==topic].reset_index(drop=True)
    rdf['P_id']=pid
    rdf['labels']=[i+label+1 for i in rdf.index.values]
    rdf=rdf.drop(columns=['Topic'])
    rdf=rdf.rename(columns={"Subtopic":"title"})
    #print(rdf)
    res_df=res_df.append(rdf).reset_index(drop=True)

print(res_df['labels'])
    
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
res_df.to_csv('transformed_impact_quality.csv')