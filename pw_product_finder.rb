require 'docx'

puts "Enter the product name to search for in the docx files: .. "
@product = gets.chomp
counter = 0
Dir.glob('C:/pw_articles/*.docx') do |docx_file|

    #puts docx_file
    d = Docx::Document.open(docx_file).to_s
    
    if d.include? @product
    	#puts d
    	puts "\nFound in: " + docx_file
    	counter += 1
    end
    
      
end

puts "\n" + counter.to_s + " files total containing this product name '#{@product}'."