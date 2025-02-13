from docx import Document
import os
import win32com.client
from src.services.config import *
import pandas as pd
from src.services.chatgpt import *
import time

def replace_paragraph_text(paragraph, old_text, new_text):
    print("A")
    for run in paragraph.runs:
        print("B")
        # print(run.text)
        if old_text in run.text:
            print("C")
            run.text = run.text.replace(old_text, new_text)
            print("D")

def replace_full_name(paragraphs, full_name):
    full_name_paragraph = paragraphs[full_name_paragraph_id]
    replace_paragraph_text(full_name_paragraph, "[Full Name]", full_name)
    
def replace_assessor_and_organization(paragraphs, assessor_name, organization):
    assessor_name_paragraph = paragraphs[assessor_name_paragraph_id]
    replace_paragraph_text(assessor_name_paragraph, "Assessor Name", assessor_name)
    replace_paragraph_text(assessor_name_paragraph, "Organization", organization)
    
def replace_background(paragraphs, background):
    background_paragraph = paragraphs[background_paragraph_id]
    replace_paragraph_text(background_paragraph, "background", background)
    
def fill_score(scores, output_doc_path):
    current_dir = os.getcwd()
    score_doc_path = os.path.join(current_dir, output_doc_path)
    # score_doc_path = "src/files/result/result.docx"
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    score_doc = word.Documents.Open(score_doc_path)
    shape_count = 0
    for shape in score_doc.Shapes:
        if shape.Type == 17:
            score = scores.loc[scores['id'] == circle_list[shape_count], 'score'].values[0]
            if " " in circle_list[shape_count]:
                shape.TextFrame.TextRange.Text = str(score)
            else:
                shape.TextFrame.TextRange.Text = int(score)
            shape_count += 1
    new_score_doc_path = os.path.join(current_dir, output_doc_path)
    score_doc.SaveAs(new_score_doc_path)
    score_doc.Close()
    word.Quit()
    
def replace_strategic_partner(paragraphs, strategic_partner):
    strategic_partner_paragraph = paragraphs[strategic_partner_paragraph_id]
    replace_paragraph_text(strategic_partner_paragraph, "strategic partner", strategic_partner)
    
def replace_innovate(paragraphs, innovate):
    innovate_paragraph = paragraphs[innovate_paragraph_id]
    replace_paragraph_text(innovate_paragraph, "innovate", innovate)
    
def replace_connect(paragraphs, connect):
    connect_paragraph = paragraphs[connect_paragraph_id]
    replace_paragraph_text(connect_paragraph, "connect", connect)
    
def replace_simplify(paragraphs, simplify):
    simplify_paragraph = paragraphs[simplify_paragraph_id]
    replace_paragraph_text(simplify_paragraph, "simplify", simplify)
    
def replace_talent_enabler(paragraphs, talent_enabler):
    talent_enabler_paragraph = paragraphs[talent_enabler_paragraph_id]
    replace_paragraph_text(talent_enabler_paragraph, "talent enabler", talent_enabler)
    
def replace_coach(paragraphs, coach):
    coach_paragraph = paragraphs[coach_paragraph_id]
    replace_paragraph_text(coach_paragraph, "coach", coach)
    
def replace_empower(paragraphs, empower):
    empower_paragraph = paragraphs[empower_paragraph_id]
    replace_paragraph_text(empower_paragraph, "empower", empower)
    
def replace_elevate(paragraphs, elevate):
    elevate_paragraph = paragraphs[elevate_paragraph_id]
    replace_paragraph_text(elevate_paragraph, "elevate", elevate)
    
def replace_agile_executor(paragraphs, agile_executor):
    agile_executor_paragraph = paragraphs[agile_executor_paragraph_id]
    replace_paragraph_text(agile_executor_paragraph, "agile executor", agile_executor)
    
def replace_deliver(paragraphs, deliver):
    deliver_paragraph = paragraphs[deliver_paragraph_id]
    replace_paragraph_text(deliver_paragraph, "deliver", deliver)
    
def replace_adapt(paragraphs, adapt):
    adapt_paragraph = paragraphs[adapt_paragraph_id]
    replace_paragraph_text(adapt_paragraph, "adapt", adapt)
    
def replace_pioneer(paragraphs, pioneer):
    pioneer_paragraph = paragraphs[pioneer_paragraph_id]
    replace_paragraph_text(pioneer_paragraph, "pioneer", pioneer)
    
def replace_resilient_steward(paragraphs, resilient_steward):
    resilient_steward_paragraph = paragraphs[resilient_steward_paragraph_id]
    replace_paragraph_text(resilient_steward_paragraph, "resilient steward", resilient_steward)
    
def replace_serve(paragraphs, serve):
    serve_paragraph = paragraphs[serve_paragraph_id]
    replace_paragraph_text(serve_paragraph, "serve", serve)
    
def replace_inspire(paragraphs, inspire):
    inspire_paragraph = paragraphs[inspire_paragraph_id]
    replace_paragraph_text(inspire_paragraph, "inspire", inspire)
    
def replace_sustain(paragraphs, sustain):
    sustain_paragraph = paragraphs[sustain_paragraph_id]
    replace_paragraph_text(sustain_paragraph, "sustain", sustain)
    
def replace_recommendation(paragraphs, transcript, scores):
    for performance in performance_list:
        print(performance)
        if scores.loc[scores['id'] == performance, 'score'].values[0] >= 7:
            print(0)
            strengthen_paragraph = paragraphs[globals()[f"{performance}_strengthen_paragraph_id"]]
            print(1)
            replace_paragraph_text(strengthen_paragraph, performance, globals()[f"get_{performance}_strengthen"](transcript))
            print(2)
            improve_paragraph = paragraphs[globals()[f"{performance}_improvement_paragraph_id"]]
            print(3)
            p = improve_paragraph._element
            print(4)
            p.getparent().remove(p)
            print(5)
            # time.sleep(30)
        else:
            print(0)
            improve_paragraph = paragraphs[globals()[f"{performance}_improvement_paragraph_id"]]
            print(1)
            replace_paragraph_text(improve_paragraph, performance, globals()[f"get_{performance}_improve"](transcript))
            print(2)
            strengthen_paragraph = paragraphs[globals()[f"{performance}_strengthen_paragraph_id"]]
            print(3)
            p = strengthen_paragraph._element
            print(4)
            p.getparent().remove(p)
            print(5)
            # time.sleep(30)
def replace_next_steps(paragraphs, next_steps):
    next_steps_paragraph = paragraphs[next_steps_paragraph_id]
    replace_paragraph_text(next_steps_paragraph, "Next steps", next_steps)

        