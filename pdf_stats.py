# If there are PDF files, show PDF file statistics
    
    if 'pdf' in file_types:
        # st.subheader("PDF File Audit:")
        pdf_file_stats = analyze_pdf_files(temp_directory)

        if pdf_file_stats:  # Check if the result is not None
            
            
            # Create lists for file names, number of pages, and file sizes
            file_names = []
            num_pages = []
            file_sizes = []

            for pdf_file, stats in pdf_file_stats.items():
                print(pdf_file,stats)
                file_names.append(pdf_file)
                num_pages.append(stats['Num Pages'])
                file_sizes.append(stats['File Size (MB)'])

            # Convert lists to DataFrame
            df_pdf_stats = pd.DataFrame({'File Name': file_names, 'Number of Pages': num_pages, 'File Size (MB)': file_sizes})

            # Display the DataFrame
            st.dataframe(df_pdf_stats)

            # Remove temporary files related to PDFs
            for pdf_file in pdf_file_stats.keys():
                file_path = os.path.join(temp_directory, pdf_file)
                os.remove(file_path)
    