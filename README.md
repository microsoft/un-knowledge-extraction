# Automatic Information Extraction and Knowledge Elicitation for United Nations Documents

#### Context: 
The processing of considerable and rapidly growing amount of information within UN system is left to the very limited human capacities. The UN system produces a substantial amount of information that, if effectively mobilized, could greatly enhance the effectiveness and efficiency of the UN system.

#### Goal: 
The goal is to pilot Microsoft Cognitive Services to unlock the strategic value of UN unstructured content by building on AI and semantic technologies. The idea is to showcase the innovative smart services that natural language processing and machine learning to effectively support policy and decision making, coordination, synergies and accountability.

#### Data: 
UN General Assembly Resolutions (English only) between 2009 and 2018. In total 3138 resolution files in pdf format.

#### Data Reference:
pre-trained word2vec embeddings trained on part of Google News dataset (about 100 billion words): https://code.google.com/archive/p/word2vec/

#### Deliverables:
##### Resolution Level:
	Resolution File Name
	Resolution Session 
	Resolution Agenda Item
	Resolution Number
	Resolution Title
	Resolution Adoption Date/Month/Year

##### Paragraph Level:
	Paragraph Type
	First Action Verb
	Key Terms
	Referenced Resolutions
	Referenced Resolution Dates
	Sustainable Development Goals (SDG), Targets, and Indicators
	Country
	Organization Names



# Setup

1. Install requirements
    
    This code use python 3.7

     ```
     pip install -r requirements.txt
     
     ```

2. Run Scripts

 	a. Run the following file for extracting resolution level information: [knowledge_extraction_resolution_level.py](https://github.com/microsoft/UN-Knowledge-Extraction/blob/main/knowledge_extraction_resolution_level.py)
    
      ```
      python knowledge_extraction_resolution_level.py
      
      ```
 	b. Run the following file for extracting paragraph level information: [knowledge_extraction_paragraph_level.py](https://github.com/microsoft/UN-Knowledge-Extraction/blob/main/knowledge_extraction_paragraph_level.py)
    
      ```
      python knowledge_extraction_paragraph_level.py
      
      ```

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
