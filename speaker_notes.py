from pptx import Presentation


def extract_speaker_notes(presentation_path):
    presentation = Presentation(presentation_path)
    speaker_notes = []

    for slide in presentation.slides:
        for shape in slide.notes_slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        speaker_notes.append(run.text)

    return speaker_notes

if __name__ == "__main__":
    presentation_path = r"C:\Users\Joshua.Tiffany\Downloads\Intern Workshop - Financial Literacy 06.21.23.pptx"
    notes = extract_speaker_notes(presentation_path)

    for note in notes:
        print(note)