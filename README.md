# AI-generated Argon-Style Presentations
Note that because of the way .pptx files get edited, this application **only works on Windows**.

The current generation is still basic, manual editing is still required on the generated presentation but it gives a reasonable starting point.

## Current Limitations
- The slide context is limited.
- Ignores the "various" category from the slide library.
- Does not generate chart, figures or images.
- The slide library is different.
- Things will improve in the future (probably).

## Setup
```bash
python -m venv .venv
venv\Scripts\activate.bat  # Command prompt
venv\Scripts\Activate.ps1  # Powershell
pip install -r requirements.txt
```

## Instructions
1. Create an empty `.env` file at the root of the directory.
2. Run `python gui.py`, it should bring up an app UI.
3. Go into settings, choose your model and add your OpenAI API keys. (Only mini models are available because the context is so big it exceeds the token limit of better models)
4. Save and exit.
5. Select a .docx file that you want to make a presentation of.
6. Click generate.
7. Click Open Output to open the generated presentation.

If you don't like the style or tone of the generation, add some additional prompts in the settings.

## Mechanics
The way this thing works is asking ChatGPT to summarise the document into a specific JSON format. 

It knows the slides from the information in `context/slide_context.json`, and then it generates a response JSON with the slots filled out. After a successful generation, you can view the response at `output/response.txt`.

Then, the `Slides_library Argon & Co.pptx` is read by the code and specific slides will be copied to a new presentation. `SlideBuilder.py` will edit these copied slides and save it result to `ōutput/output.pptx`. There is a backup slide library because sometimes if the program crashes during reading or editing, it will literally eat the slides that have been read (i.e. slides just disappear).

Note that the `Slides_library Argon & Co.pptx` is actually slightly different compared to the one in Qua'Li, because I had to rename some element IDs to make them unique, but the slides are the exact same.

If you don't like the slide style or content it generated. Go into the `output/response.txt` and edit the fields.
- `slide_index` is the slide number in the `Slides_library Argon & Co.pptx`.
- `slots` are the available fields for editing for that specific slide. Check `SlideBuilder.py` and use selection pane feature in PowerPoint to see which element is which.
- Once the response is edited, click on regenerate to produce a new `output.pptx`. Note that this does **NOT** prompt the AI again, the presentation is generated purely on the .txt file. Maybe in the future I will add a verbal feedback mechanic.

Or just edit the `output.pptx`.
