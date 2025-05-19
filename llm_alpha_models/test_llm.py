from transformers import pipeline

# Initialize a text generation pipeline with GPT-Neo 2.7B
generator = pipeline('text-generation', model='EleutherAI/gpt-neo-2.7B')

# Define the prompt
prompt = "do you know how to analyze an excel spreadsheet with complex formulas?"

# Generate a response
response = generator(prompt, max_length=50, num_return_sequences=1)

# Print the response
print("LLM response:", response[0]['generated_text']) 