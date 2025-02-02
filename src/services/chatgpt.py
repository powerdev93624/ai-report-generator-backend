from openai import OpenAI
import os
import dotenv
import httpx

dotenv.load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if os.environ.get("APP_ENV") == 'development':
    client = OpenAI(
        api_key=OPENAI_API_KEY,
        http_client=httpx.Client(
            proxy=os.getenv("OPENAI_PROXY"),
            transport=httpx.HTTPTransport(local_address="0.0.0.0"),
        ),
    )
else:
    client = OpenAI(
        api_key=OPENAI_API_KEY,
    )



def get_response(prompt):
    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an interview transcript analyzer. You should analyze the interview transcript."},
            {"role": "user", "content": prompt}
        ],
    )

    return completion.choices[0].message.content

def get_background(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200 word summary of the individual's background, education, and career trajectory, tailored to the executive being assessed. 
    
    Sample Answer:
    Ahmed Khaled Al Harbi began his career with a formative experience studying in the United States, which significantly shaped his independence and self-reliance. During his time abroad, he developed resilience and adaptability, navigating cultural differences and building a foundation of confidence. Upon returning to the UAE, Ahmed pursued a degree in Electronics and Computers at the Higher Colleges of Technology, excelling in hands-on technical projects such as building and configuring computer networks.

    Ahmed’s professional career spans over 15 years at ABC Corporation, where he has worked primarily in the IT Networking Division. His roles have included both individual contributor and project leadership responsibilities, with a focus on designing and implementing network infrastructures. Known for his calm demeanor and collaborative approach, Ahmed has demonstrated technical expertise and reliability, often working during weekends and late hours to ensure system integrity. Recently, Ahmed has shown enthusiasm for cross-departmental initiatives, participating in committees that broaden his perspective and impact beyond IT.

    """
    return get_response(prompt)

def get_strategic_partner(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 50-100 word summary of the executive's performance in this domain. 
    
    Domain Definition: STC's STAR Leaders demonstrate progressive long-term thinking and a growth mindset, build strong partnerships to foster collaboration, and simplify complexity to align priorities and maximize stakeholder value.
    
    Sample Answer:
    Ahmed demonstrates a strong technical mindset and an ability to build effective working relationships with colleagues and vendors. He is collaborative and open to learning, which supports his ability to connect with others. However, he could enhance his ability to drive strategic initiatives and simplify complex problems independently.

    """
    return get_response(prompt)

def get_innovate(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200 word detailed narrative summarizing the executive's behaviors and experiences related to this competency. 
    
    Domain Definition: Demonstrates a growth mindset and progressive long-term thinking by challenging the status quo to identify and promote innovative strategic initiatives.
    
    Sample Answer:
    Ahmed’s technical expertise allows him to identify practical solutions, such as redesigning network protocols to improve efficiency. For instance, during a critical systems review, Ahmed identified several inefficiencies in the existing network setup and proposed a phased improvement plan that immediately boosted operational speed by 20%. His methodical approach to resolving issues has been instrumental in maintaining system reliability, even in high-pressure situations. However, Ahmed tends to focus on incremental enhancements rather than challenging existing frameworks. For example, during a project to modernize IT infrastructure, he opted for minor adjustments rather than advocating for a transformative overhaul. While this ensured short-term stability, it limited the organization’s ability to fully leverage emerging technologies. Developing his ability to think beyond immediate solutions and embrace bold, strategic innovation will greatly enhance his leadership impact. Ahmed has expressed interest in exploring innovative strategies, indicating a willingness to grow in this area.

    """
    return get_response(prompt)

def get_connect(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200 word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Develops and maintains strong relationships with strategic stakeholders (internal and external partners / government / customers) by actively enabling collaboration across the enterprise.
    
    Sample Answer:
    Ahmed excels at building strong interpersonal connections, both within his team and with external vendors. For example, he successfully negotiated a cost-saving agreement with a key supplier by leveraging mutual respect and clear communication, which led to a 15% reduction in overall expenses for critical infrastructure upgrades. Ahmed's ability to communicate effectively and empathetically has earned him a reputation as a trusted collaborator, particularly when working across departments. His participation in cross-departmental committees has broadened his organizational perspective, allowing him to align IT objectives with broader corporate strategies. One notable example involved his leadership in a cross-functional task force to streamline procurement processes, where his contributions ensured smoother coordination and timely execution of projects. While Ahmed’s relational skills are strong, there is room to align these relationships more strategically with organizational goals. Developing a more proactive approach to leveraging these partnerships for long-term strategic initiatives could further amplify his impact as a leader at STC.

    """
    return get_response(prompt)

def get_simplify(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200 word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Aligns the enterprise on highest-priority efforts that maximize stakeholder value by aggressively reducing complexity and bureaucracy.
    
    Sample Answer:
    Ahmed is adept at breaking down technical problems into manageable tasks, enabling his team to implement solutions efficiently. For instance, during a recent network outage, he swiftly identified the root cause and delegated responsibilities to restore operations within hours, ensuring minimal disruption to business operations. His hands-on approach to troubleshooting and clear communication kept his team focused and motivated under pressure. Ahmed’s ability to dissect complex issues is also evident in his work on upgrading IT infrastructure, where he reorganized workflows to reduce redundancies and improve system performance by 15%. 
    However, Ahmed’s focus is often on immediate technical outcomes rather than addressing systemic complexities that could prevent future issues. For example, while he excels in managing short-term crises, he could further develop strategies to simplify and streamline processes across the organization. By adopting a broader perspective on organizational priorities and proactively reducing bureaucracy, Ahmed could significantly enhance his strategic contributions to STC. His willingness to collaborate and learn positions him well to grow in this area, with targeted mentorship and leadership training.
    """
    return get_response(prompt)

def get_talent_enabler(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 50-100 word summary of the executive's performance in this domain.
    
    Domain Definition: STC's STAR Leaders actively coach, develop, and engage talent, empower high-performing teams by cultivating trust and ownership, and elevate performance by inspiring excellence and enabling individuals and teams to achieve exceptional results.
    
    Sample Answer:
    Ahmed is collaborative and works well within a team but prefers to remain a contributor rather than a leader. He could develop his skills in coaching and empowering others, particularly in leadership and mentorship roles.
    """
    return get_response(prompt)

def get_coach(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200 word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Actively coaches and develops talent by fostering adaptability and preparing them to lead effectively in an ever-evolving environment.
    
    Sample Answer:
    Ahmed has a reputation for sharing his technical expertise informally, particularly with junior colleagues. For instance, he has been instrumental in mentoring several new hires within the IT Networking Division, helping them troubleshoot complex systems and learn the nuances of network configuration. However, Ahmed has not yet established a structured coaching framework, which could amplify his impact. During a critical system migration project, Ahmed worked closely with a team member struggling to adapt to the new protocols, guiding them step-by-step to understand the requirements and succeed in their role. This experience highlighted his patience and ability to foster growth. Despite these contributions, Ahmed tends to focus on his individual responsibilities rather than adopting a proactive coaching role. Developing a consistent approach to coaching, with formal check-ins and feedback sessions, could help him nurture future leaders within the organization.
    """
    return get_response(prompt)

def get_empower(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200-word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Builds high-performing teams by cultivating trust, delegating authority, and creating an environment where team members take ownership, collaborate effectively, and maximize their potential.
    
    Sample Answer:
    Ahmed has demonstrated a collaborative spirit, often serving as a key team player in complex technical projects. For example, during a network optimization initiative, Ahmed delegated certain tasks to junior staff, empowering them to take ownership while providing necessary support. However, his preference for individual contribution sometimes limits his ability to fully embrace delegation and team empowerment. In another instance, Ahmed worked with a cross-functional team to streamline processes, where his ability to create a trusting and inclusive environment ensured the project stayed on track. Despite these successes, Ahmed occasionally hesitates to entrust critical responsibilities to others, which can slow team growth and innovation. By learning to delegate more effectively and cultivating ownership among team members, Ahmed can unlock higher levels of team performance and expand his leadership capacity.
    """
    return get_response(prompt)

def get_elevate(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200-word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Accelerates peak performance by setting high expectations, inspiring those around them, and enabling individuals and teams to achieve exceptional results through continuous support and constructive feedback.
    
    Sample Answer:
    Ahmed consistently strives for excellence in his work, setting a high standard for quality and attention to detail. For instance, during a company-wide system upgrade, his thorough planning and execution inspired the team to meet tight deadlines without compromising quality. While Ahmed provides constructive feedback in informal settings, he has not yet established a structured approach to performance reviews. During a particularly challenging project, Ahmed’s encouragement and troubleshooting support motivated a team member to overcome significant obstacles and deliver results exceeding expectations. Despite these strengths, Ahmed could enhance his leadership by setting clearer performance expectations and offering more frequent, structured feedback sessions. By doing so, he can foster a culture of continuous improvement and inspire his team to achieve even greater success.
    """
    return get_response(prompt)

def get_agile_executor(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 50-100 word summary of the executive's performance in this domain.
    
    Domain Definition: STC's STAR Leaders deliver results by fostering accountability and tracking outcomes, adapt to change by navigating complexity with agility, and pioneer digital transformation by championing emerging technologies and innovative solutions.
    
    Sample Answer:
    Ahmed excels as an Agile Executor, showcasing a disciplined focus on delivering results, adaptability in ambiguous situations, and a proactive approach to leveraging emerging technologies. His ability to manage high-pressure projects and lead teams through change is commendable, though there is room to enhance strategic alignment and broader organizational integration of digital initiatives. Strengthening these areas will amplify Ahmed’s already impressive contributions to driving results and innovation.
    """
    return get_response(prompt)

def get_deliver(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200-word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Creates a culture of accountability by aligning goals, clarifying key metrics, tracking outcomes, and maintaining a disciplined focus on achieving results that advance strategic priorities.
    
    Sample Answer:
    Ahmed has a strong track record of delivering results, particularly in high-pressure situations. For example, during a mission-critical server migration, Ahmed meticulously aligned his team’s efforts with project milestones, ensuring that key deliverables were met on time. He is highly disciplined in tracking outcomes and maintaining focus, which has earned him the trust of his managers and peers. However, Ahmed’s approach tends to focus on technical execution rather than strategic alignment. During a recent initiative to improve network security, he successfully met the technical goals but could have engaged more with stakeholders to align the project outcomes with broader organizational priorities. By developing his ability to connect individual project goals with the company’s strategic vision, Ahmed can further enhance his effectiveness as a results-driven leader.
    """
    return get_response(prompt)

def get_adapt(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200-word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Leads change effectively by embracing ambiguity, removing barriers to progress, and guiding teams to navigate complex challenges with agility and confidence.
    
    Sample Answer:
    Ahmed demonstrates a natural ability to remain composed and resourceful in the face of ambiguity. During a major system transition involving unforeseen technical issues, Ahmed quickly assessed the situation and rallied his team to troubleshoot effectively, ensuring minimal disruption to operations. His proactive problem-solving approach allowed him to address uncertainties while keeping his team focused. In another example, Ahmed played a key role in a project involving shifting organizational priorities, where his ability to adapt to new directives without losing momentum was critical to the project’s success. However, Ahmed’s focus on immediate solutions occasionally limits his ability to address larger, systemic challenges. By building on his adaptability and honing his skills in removing barriers to progress, Ahmed can further strengthen his leadership capacity in driving change.
    """
    return get_response(prompt)

def get_pioneer(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200-word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Drives digital transformation and industry leadership by understanding emerging technologies, championing their implementation, and prioritizing digital solutions to sustain a strong competitive edge.
    
    Sample Answer:
    Ahmed has consistently demonstrated a keen interest in leveraging emerging technologies to enhance organizational performance. For instance, he spearheaded the adoption of advanced network analytics tools that significantly improved system efficiency and reliability. His ability to stay updated on technological trends enables him to identify and advocate for tools that align with STC’s strategic goals. Additionally, Ahmed has contributed to workshops aimed at fostering innovation within his team, encouraging them to explore digital solutions for longstanding challenges. However, while Ahmed excels in identifying and implementing technical solutions, he could improve in championing digital transformation on a broader organizational scale. By developing a more strategic focus on integrating these technologies into STC’s long-term vision, Ahmed can elevate his impact as a pioneer in the digital domain.
    """
    return get_response(prompt)

def get_resilient_steward(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 50-100 word summary of the executive's performance in this domain..
    
    Domain Definition: STC's STAR Leaders serve stakeholders with humility by embracing new  perspectives, inspire others with optimism and integrity, and sustaining long-term success by embodying grit and resilience.
    
    Sample Answer:
    Ahmed exemplifies resilience and humility, fostering trust and collaboration across teams while remaining steadfast in the face of challenges. His ability to incorporate diverse perspectives and maintain composure under pressure has earned widespread respect. While Ahmed’s leadership is defined by integrity and perseverance, further developing his ability to inspire and instill resilience in others will enhance his impact, enabling him to lead with greater influence and drive long-term success at STC.
    """
    return get_response(prompt)

def get_serve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200 word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Embodies humility by setting aside ego, staying open to diverse ideas, and remaining flexible when faced with new perspectives.
    
    Sample Answer:
    Ahmed collaborates effectively across teams and demonstrates a consistent openness to diverse perspectives. For instance, during a cross-departmental initiative to redesign IT workflows, Ahmed actively sought input from colleagues outside of his technical expertise, incorporating their feedback into the final solution. This approach not only strengthened the project outcomes but also earned him respect for his humility and willingness to listen. In another instance, Ahmed’s adaptability was tested when faced with conflicting stakeholder priorities during a critical system upgrade. By setting aside personal preferences and focusing on collaborative solutions, he ensured the project’s success while maintaining team harmony. However, Ahmed could further develop his ability to take initiative in proposing solutions. Encouraging him to lead discussions and actively challenge the status quo will enhance his contributions and allow him to make a broader impact as a leader at STC.
    """
    return get_response(prompt)

def get_inspire(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200 word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Leads with optimism and integrity by fostering trust, living STC values, demonstrating a positive outlook, and setting an example that motivates others to overcome challenges with confidence and purpose.
    
    Sample Answer:
    Ahmed’s calm demeanor and integrity have consistently fostered trust among his colleagues. For example, during a high-pressure IT infrastructure upgrade, Ahmed remained composed, rallying his team to troubleshoot issues collaboratively. In one instance, when a critical system failure threatened to delay operations, Ahmed's level-headedness not only restored order but also motivated his team to push through a challenging night shift. Additionally, Ahmed’s transparency and ethical decision-making have earned him respect across departments. His regular efforts to communicate honestly about challenges and decisions create an environment of trust, ensuring stakeholders are aligned and confident in his leadership. However, Ahmed often focuses on task execution over actively inspiring his team. For example, while his technical proficiency is appreciated, he has yet to develop a compelling vision that galvanizes his team beyond immediate goals. This tendency could limit his influence as he steps into roles requiring broader leadership and inspiration.
    """
    return get_response(prompt)

def get_sustain(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 150-200-word detailed narrative summarizing the executive's behaviors and experiences related to this competency.
    
    Domain Definition: Demonstrates grit and resilience by persevering through challenges, maintaining focus on long-term goals, and encouraging others to stay committed in the face of adversity.
    
    Sample Answer:
    Ahmed’s resilience has been a hallmark of his career, particularly during high-pressure situations. For example, during a prolonged system outage that required multiple overnight shifts, Ahmed maintained his focus and encouraged his team to persevere until the issue was resolved. His determination and calm leadership kept morale high, even under challenging circumstances. Additionally, Ahmed’s commitment to long-term goals is evident in his role as a key contributor to the rollout of a multi-year IT modernization plan, where he consistently met deadlines and ensured project milestones were achieved. Despite his strengths, Ahmed could further enhance his leadership by encouraging his team to adopt a similarly resilient mindset, particularly when faced with setbacks. By sharing his strategies for overcoming challenges and fostering a culture of perseverance, Ahmed can inspire his colleagues to achieve even greater results.

    """
    return get_response(prompt)

def get_innovate_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.    
    
    Domain Definition: Demonstrates a growth mindset and progressive long-term thinking by challenging the status quo to identify and promote innovative strategic initiatives.
    
    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_connect_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Develops and maintains strong relationships with strategic stakeholders (internal and external partners / government / customers) by actively enabling collaboration across the enterprise.
    
    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_simplify_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Aligns the enterprise on highest-priority efforts that maximize stakeholder value by aggressively reducing complexity and bureaucracy.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_coach_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Actively coaches and develops talent by fostering adaptability and preparing them to lead effectively in an ever-evolving environment.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_empower_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Builds high-performing teams by cultivating trust, delegating authority, and creating an environment where team members take ownership, collaborate effectively, and maximize their potential.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_elevate_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Accelerates peak performance by setting high expectations, inspiring those around them, and enabling individuals and teams to achieve exceptional results through continuous support and constructive feedback.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_deliver_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Creates a culture of accountability by aligning goals, clarifying key metrics, tracking outcomes, and maintaining a disciplined focus on achieving results that advance strategic priorities.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_adapt_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Leads change effectively by embracing ambiguity, removing barriers to progress, and guiding teams to navigate complex challenges with agility and confidence.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_pioneer_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Drives digital transformation and industry leadership by understanding emerging technologies, championing their implementation, and prioritizing digital solutions to sustain a strong competitive edge.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_serve_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Embodies humility by setting aside ego, staying open to diverse ideas, and remaining flexible when faced with new perspectives.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_inspire_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Leads with optimism and integrity by fostering trust, living STC values, demonstrating a positive outlook, and setting an example that motivates others to overcome challenges with confidence and purpose.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_sustain_strengthen(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation about executive's strength in this domain.
    
    Domain Definition: Demonstrates grit and resilience by persevering through challenges, maintaining focus on long-term goals, and encouraging others to stay committed in the face of adversity.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_innovate_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.  
    
    Domain Definition: Demonstrates a growth mindset and progressive long-term thinking by challenging the status quo to identify and promote innovative strategic initiatives.
    
    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_connect_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Develops and maintains strong relationships with strategic stakeholders (internal and external partners / government / customers) by actively enabling collaboration across the enterprise.
    
    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_simplify_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Aligns the enterprise on highest-priority efforts that maximize stakeholder value by aggressively reducing complexity and bureaucracy.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_coach_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Actively coaches and develops talent by fostering adaptability and preparing them to lead effectively in an ever-evolving environment.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_empower_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Builds high-performing teams by cultivating trust, delegating authority, and creating an environment where team members take ownership, collaborate effectively, and maximize their potential.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_elevate_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Accelerates peak performance by setting high expectations, inspiring those around them, and enabling individuals and teams to achieve exceptional results through continuous support and constructive feedback.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_deliver_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Creates a culture of accountability by aligning goals, clarifying key metrics, tracking outcomes, and maintaining a disciplined focus on achieving results that advance strategic priorities.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_adapt_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Leads change effectively by embracing ambiguity, removing barriers to progress, and guiding teams to navigate complex challenges with agility and confidence.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_pioneer_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Drives digital transformation and industry leadership by understanding emerging technologies, championing their implementation, and prioritizing digital solutions to sustain a strong competitive edge.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_serve_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Embodies humility by setting aside ego, staying open to diverse ideas, and remaining flexible when faced with new perspectives.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_inspire_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Leads with optimism and integrity by fostering trust, living STC values, demonstrating a positive outlook, and setting an example that motivates others to overcome challenges with confidence and purpose.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_sustain_improve(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide one-sentence of 15-20 words recommendation to improve executive's performance in this domain.
    
    Domain Definition: Demonstrates grit and resilience by persevering through challenges, maintaining focus on long-term goals, and encouraging others to stay committed in the face of adversity.

    Sample Answer: Ahmed excels at building strong relationships with stakeholders, which is critical for cross-functional collaboration.
    """
    return get_response(prompt)

def get_next_steps(transcript):
    prompt = f"""
    Below is the interview transcript.
    
    =====================Transcript========================
    {transcript}
    =======================================================
    
    Please provide 90-120 words completed next steps for the executive as below format.
    Please use real first name instead of executive.
    
    To build on the executive’s strengths and address their development areas, we recommend … [ insert recommendations ]. Key next steps include [ insert next steps ]. Additionally, [insert First Name] should set personal development goals aligned with the STC STAR competencies and regularly review progress with their manager. By taking these steps, [insert First Name] will be well-positioned to grow as a leader and make a meaningful impact on STC’s long-term success. 
    Domain Definition: Demonstrates grit and resilience by persevering through challenges, maintaining focus on long-term goals, and encouraging others to stay committed in the face of adversity.

    Sample Answer: To build on the executive’s strengths and address their development areas, we recommend engaging in a structured leadership development program tailored to the STC High-Potential Competency Model. Key next steps include mentoring from senior leaders to enhance strategic thinking and coaching capabilities, participation in targeted workshops on innovation and complexity reduction, and involvement in cross-functional projects to broaden their organizational impact. Additionally, the executive should set personal development goals aligned with the STC competencies and regularly review progress with their manager. By taking these steps, the executive will be well-positioned to grow as a leader and make a meaningful impact on STC’s long-term success.
    """
    return get_response(prompt)
