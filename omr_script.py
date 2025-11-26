import os
import fitz  # PyMuPDF
from PIL import Image
import cv2
import numpy as np
import pandas as pd  # Import pandas for Excel file creation

def process_omr(pdf_path, images_folder_path, output_folder, answer_key, num_of_questions, height_per_question=12):
    """ Process OMR PDF, convert pages to images, process them, and calculate marks based on provided answer key. """

    # Step 1: Convert PDF to images
    def pdf_to_images(pdf_path, output_folder, img_size=(342, 486)):
        """ Convert each page of the PDF to images and save them in the output folder. """
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        pdf_document = fitz.open(pdf_path)

        for i in range(pdf_document.page_count):
            page = pdf_document.load_page(i)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_resized = img.resize(img_size, Image.Resampling.LANCZOS)

            output_path = os.path.join(output_folder, f"page_{i + 1}.png")
            img_resized.save(output_path)
            print(f"Saved page {i + 1} to {output_path}")

        pdf_document.close()

    # Step 2: Crop the answer region from the OMR sheet
    def crop_answer_region(image, save_cropped_image_path=None):
        """ Crops the answer region from the OMR sheet based on provided offsets and saves it as an image. """
        height, width = image.shape[:2]

        # Define offsets
        top_offset = 330
        bottom_offset = 30
        left_offset = 17
        right_offset = 19

        # Calculate cropping coordinates
        x1 = left_offset
        y1 = top_offset
        x2 = width - right_offset
        y2 = height - bottom_offset

        answer_region = image[y1:y2, x1:x2]

        if answer_region.size == 0:
            print("Error: The answer region is empty.")
            return None

        if save_cropped_image_path:
            cv2.imwrite(save_cropped_image_path, answer_region)
            print(f"Cropped answer region saved to {save_cropped_image_path}")

        return answer_region

    # Step 3: Divide the cropped image into 3 parts
    def divide_cropped_image(image, output_folder):
        """ Divides the cropped image into 3 parts. """
        height, width = image.shape[:2]

        part1_width = 123
        part2_width = 62
        part3_width = 121

        if part1_width + part2_width + part3_width != width:
            print("Warning: The sum of the parts' widths doesn't match the image width. Adjusting the last part.")
            part3_width = width - part1_width - part2_width

        part1 = image[:, :part1_width]
        part2 = image[:, part1_width:part1_width + part2_width]
        part3 = image[:, part1_width + part2_width:]

        cv2.imwrite(os.path.join(output_folder, "part1.png"), part1)
        cv2.imwrite(os.path.join(output_folder, "part2.png"), part2)
        cv2.imwrite(os.path.join(output_folder, "part3.png"), part3)

        print("Divided images saved successfully!")

    # Step 4: Divide image into individual question images
    def divide_image_into_questions(image, output_folder, start_question, end_question, height_per_question, fixed_questions=4):
        """ Divides the image into individual question images with variable heights. """
        height, width = image.shape[:2]

        num_questions = end_question - start_question + 1
        total_height_fixed = fixed_questions * height_per_question
        remaining_height = height - total_height_fixed
        remaining_questions = num_questions - fixed_questions

        if remaining_questions > 0:
            height_per_remaining_question = remaining_height // remaining_questions
        else:
            height_per_remaining_question = height_per_question

        y_offset = 0

        for Q_idx in range(num_questions):
            if Q_idx < fixed_questions:
                y1 = y_offset
                y2 = y1 + height_per_question
                y_offset = y2
            else:
                y1 = y_offset
                y2 = y1 + height_per_remaining_question
                y_offset = y2

            if y2 > height:
                y2 = height

            Q_image = image[y1:y2, :]
            Q_filename = f"Q_{start_question + Q_idx}.png"
            Q_path = os.path.join(output_folder, Q_filename)
            cv2.imwrite(Q_path, Q_image)
            print(f"Saved: {Q_path}")

    # Step 5: Process each image and divide them into questions
    def process_part_image(part_image_path, output_folder, start_question, end_question, height_per_question, fixed_questions=4):
        """ Process and divide part image into individual question images. """
        image = cv2.imread(part_image_path)

        if image is None:
            print(f"Error: Could not load image at {part_image_path}")
            return

        print(f"Image dimensions: {image.shape}")

        divide_image_into_questions(image, output_folder, start_question, end_question, height_per_question, fixed_questions)

    # Step 6: Detect the darkest section
    def divide_and_detect_darkest_part(image_path, num_parts=5):
        """ Divide the image into parts and detect the darkest section. """
        image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)

        if image is None:
            print(f"Error: Could not load image at {image_path}")
            return None

        height, width = image.shape[:2]
        part_width = width // num_parts

        darkest_part = None
        darkest_value = 255
        darkest_part_index = -1

        labels = ['NA', 'A', 'B', 'C', 'D']

        for i in range(num_parts):
            if i == num_parts - 1:
                part = image[:, i * part_width:]
            else:
                part = image[:, i * part_width:(i + 1) * part_width]

            avg_intensity = np.mean(part)

            if avg_intensity < darkest_value:
                darkest_value = avg_intensity
                darkest_part = part
                darkest_part_index = i

        return labels[darkest_part_index]

    # Step 7: Process multiple images for darkest section detection
    def process_multiple_images(image_paths):
        """ Process multiple images for darkest section detection. """
        results = {}
        for idx, image_path in enumerate(image_paths):
            section = divide_and_detect_darkest_part(image_path)
            if section:
                image_name = f"Q_{idx + 1}"
                results[image_name] = (image_path, section)

        return results

    # Step 8: Compare generated answers with the answer key and calculate marks
    def compare_answers_and_calculate_marks(generated_answers, answer_key, num_of_questions):
        """ Compare the generated answers with the provided answer key and calculate marks. """
        correct_answers = 0
        Q_results = {}
        for i in range(num_of_questions):
            question = f"Q_{i + 1}"
            if question in generated_answers and generated_answers[question][1] == answer_key[i]:
                correct_answers += 1
                Q_results[question] = 1
            else:
                Q_results[question] = 0

        marks = correct_answers / num_of_questions * 100
        return correct_answers, marks, Q_results

    # Convert PDF to images
    pdf_to_images(pdf_path, images_folder_path)

    # Initialize results storage
    results_list = []

    # Process each image in the output folder
    for image_file in os.listdir(images_folder_path):
        if image_file.lower().endswith(('.png', '.jpg', '.jpeg')):
            omr_image_path = os.path.join(images_folder_path, image_file)
            print(f"Processing image: {omr_image_path}")

            # Process the OMR image
            image = cv2.imread(omr_image_path, cv2.IMREAD_COLOR)
            cropped_image_path = os.path.join(output_folder, "answer_region.png")
            cropped_image = crop_answer_region(image, save_cropped_image_path=cropped_image_path)

            if cropped_image is not None:
                divide_cropped_image(cropped_image, output_folder)

                # Process part1 and part3 for questions
                part1_image_path = os.path.join(output_folder, "part1.png")
                process_part_image(part1_image_path, output_folder, start_question=1, end_question=10, height_per_question=height_per_question)

                part3_image_path = os.path.join(output_folder, "part3.png")
                process_part_image(part3_image_path, output_folder, start_question=11, end_question=20, height_per_question=height_per_question)

                # Process multiple images for darkest detection
                image_paths = [os.path.join(output_folder, f"Q_{i}.png") for i in range(1, 21)]
                generated_answers = process_multiple_images(image_paths)

                # Compare the generated answers with the provided answer key and calculate marks
                correct_answers, marks, Q_results = compare_answers_and_calculate_marks(generated_answers, answer_key, num_of_questions)

                # Store the results
                result_entry = {
                    'image_name': image_file
                }
                result_entry.update(Q_results)
                result_entry.update({
                    'correct_answers': f"{correct_answers} out of {num_of_questions}",
                    'marks': marks
                })
                results_list.append(result_entry)

                # Print the results
                print(f"\nCorrect Answers for {image_file}: {correct_answers} out of {num_of_questions}")
                print(f"Marks for {image_file}: {marks}%\n")

    # Save results to an Excel file
    results_df = pd.DataFrame(results_list)
    results_df.to_excel(os.path.join(output_folder, 'omr_results.xlsx'), index=False)
    print("Results saved to 'omr_results.xlsx'.")

# Example usage with user input:
if __name__ == "__main__":
    pdf_path = r"D:\srm_omr\uploads\omr_sheets.pdf"
    images_folder_path = r"D:\srm_omr\images_output"
    output_folder = r"D:\srm_omr\uploads\output"
    # Input teacher name from user
    teacher_name = input("Enter the name of the teacher :")

    # Input subject from user
    subject = input("Enter the name of the subject :")

    # Input number of questions to check
    num_of_questions = int(input("Enter the number of questions to check: "))

    # Input answer key from user
    answer_key = input("Enter the answer key as comma-separated values (e.g., A,B,C,D,...): ").split(',')

    # Call the merged function
    process_omr(pdf_path, images_folder_path, output_folder, answer_key, num_of_questions)
