require 'google/apis/classroom_v1'
require 'google/apis/drive_v3'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'fileutils'
require 'parallel'
require 'byebug'
require 'Digest/md5'
require 'yaml'


# execute 'gem install google-api-client' to install all google dependencies

# format of COURSES_STUDENTS_FILE yaml-file should resemble this:

# Programmering 1:
# - John Andersson
# - David Johansson
#
# Webbutveckling 1:
# - Peter Eriksson
# - Maria Olsson

class GoogleClassRoomAssignmentDownloader
	
	COURSES_STUDENTS_FILE = "courses_students.yaml"
	OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'
	APPLICATION_NAME = 'Classroom Assignment Downloader'
	CLIENT_SECRETS_PATH = 'client_secret.json'
	CREDENTIALS_PATH = File.join(Dir.home, '.credentials', "classroom.googleapis.com-classroom-assignment-downloader.yaml")
	SCOPE = [
	Google::Apis::ClassroomV1::AUTH_CLASSROOM_COURSES_READONLY, 
	Google::Apis::ClassroomV1::AUTH_CLASSROOM_COURSEWORK_STUDENTS_READONLY, 
	Google::Apis::DriveV3::AUTH_DRIVE_READONLY,
	Google::Apis::DriveV3::AUTH_DRIVE_METADATA_READONLY
]

def initialize
	@service = initialize_service
	@courses = list_courses
	display_course_menu
	@selected_course = select_course
	@assignments = list_assignments
	display_assignments_for_selected_course
	@selected_assignment = select_assignment
	@files = list_files
	download_files	
end

def list_courses
	response = @service.list_courses(course_states:["ACTIVE"])
	response.courses.each_with_index.map {|course, index| {option: index + 1, name:course.name, id: course.id }}
end

def display_course_menu
	@courses.each { |course| puts "course: #{course[:option]}: #{course[:name]}" }
	print "Please select course: "
end

def select_course
	selection = gets.to_i		
	selected_course = @courses.select { |course| course[:option] == selection }.first
end

def list_assignments
	response = @service.list_course_works(@selected_course[:id])
	response.course_work.each_with_index.map {|work, index| {option: index + 1, name:work.title, id: work.id }}
end

def display_assignments_for_selected_course
	@assignments.each { |assignment| puts "Assignment: #{assignment[:option]}: #{assignment[:name]}" }
	print "Please select assignment: "
end

def select_assignment
	selection = gets.to_i		
	@assignments.select { |assignment| assignment[:option] == selection }.first
end

def list_files
	response = @service.get_course_work(@selected_course[:id], @selected_assignment[:id])
	files = nil
	if response.assignment
		drive_folder_id = response.assignment.student_work_folder.id
		
		@drive_service = Google::Apis::DriveV3::DriveService.new
		@drive_service.client_options.application_name = APPLICATION_NAME
		@drive_service.authorization = authorize_service
		
		# 2019-10-28 - added properties: mimeType, md5Checksum, modifiedTime
		file_metadata_fields = "files(id,name,originalFilename,sharingUser(displayName,emailAddress),owners(displayName,emailAddress),mimeType,md5Checksum,modifiedTime),nextPageToken"
		response = @drive_service.list_files(q: "'#{drive_folder_id}' in parents", fields: file_metadata_fields )		
		files = response.files
		while response.next_page_token
			response = @drive_service.list_files(page_token: response.next_page_token, q: "'#{drive_folder_id}' in parents", fields: file_metadata_fields )		
			files += response.files
		end
	end
	files = remove_duplicate_drive_files( files )
end

# @drive_service.list_files will under certain circumstances return more than revision of same binary file
# this will sort the files by email_address (unique username), filename, modified_time and remove all but
# the most recent copy of the file
def remove_duplicate_drive_files( files )
	files_new = []
	files.each do |file|
		email_address = file.sharing_user ? file.sharing_user.email_address : file.owners.first.email_address
		files_new << [ email_address, file.name, file.modified_time, file ]
	end
	# sort by email_address, filename and (descending) file's modified date
	files_sorted = files_new.sort_by {|file| [file[0], file[1], -file[2].to_time.to_i, file[3]] }

	puts ""
	files_no_duplicates = []
	i = 0
	email_address = files_sorted[i][0]
	filename = files_sorted[i][1]
	files_no_duplicates << files_sorted[0]
	while i < files_sorted.size - 1
		i += 1
		unless email_address == files_sorted[i][0] and filename == files_sorted[i][1]
			files_no_duplicates << files_sorted[i] 
		else
			puts "Ignoring duplicate file #{email_address}::#{filename}::#{files_sorted[i][2]}"
		end
		email_address = files_sorted[i][0]
		filename = files_sorted[i][1]    
	end
	files_sorted = []
	files_no_duplicates.each { |file| files_sorted << file[3] }
	files_sorted
end

# rename existing copy of a file per "basename yyyy-mm-dd hh.mm.ss.extension"
# this will typically happen to revised assignments (same filename uploaded multiple times)
def rename_existing_file( file_path )
	local_file_mtime = File.mtime( file_path )					# modified time
	local_file_mtime_str = local_file_mtime.inspect[0..-7]		# "2019-06-01 22:37:25 +0200" -> "2019-06-01 22:37:25"
	local_file_extension = File.extname( file_path )
	local_file_basename = File.basename( file_path, local_file_extension )
	user_path = File.dirname( file_path )
	file_path_previous_copy = File.join( user_path, "#{local_file_basename} (#{local_file_mtime_str.tr(':', '.')})#{local_file_extension}" )
	File.rename( file_path, file_path_previous_copy )
	return file_path_previous_copy
end

def get_safe_filepath( filename, *path )
	# replace invalid filename characters
	safe_filename = filename.gsub(/[\/\\?*:|"<>]/, '_')
	file_path = File.join( path, safe_filename )
end

# restore modified time for a downloaded file, adjust for UTC+02:00
def restore_file_time( file, file_path )
	# get file mtime -> 2019-04-24T11:12:29+00:00 (add two hours to utc offset for swedish local time)
	drive_file_mtime = file.modified_time.to_time.getlocal("+02:00")
	File.utime( drive_file_mtime, drive_file_mtime, file_path )	# utime( atime, mtime )
end

# remove duplicate file based on matching timestamps
# returns true if a duplicate was removed, false otherwise
def remove_duplicate_file_by_timestamp( file_path, file_path_previous_copy )
	if file_path_previous_copy
		previous_copy_mtime = File.mtime( file_path_previous_copy )
		current_file_mtime = File.mtime( file_path )
		# if files appear to be identical -> remove one copy
		if previous_copy_mtime.to_i == current_file_mtime.to_i			# hack to accept same date (.utc? will mismatch)
			File.delete( file_path_previous_copy )
			# puts "Deleting identical file: #{file_path_previous_copy}"
			return true
		end
	end
	return false
end

# remove duplicate file based on file size
# returns true if a duplicate was removed, false otherwise
def remove_duplicate_file_by_size( file_path, file_path_previous_copy )
	if file_path_previous_copy
		# if files appear to be identical -> remove one copy
		if File.size( file_path_previous_copy ) == File.size( file_path )
			File.delete( file_path_previous_copy )
			# puts "Deleting identical file: #{file_path_previous_copy}"
			return true
		end
	end
	return false
end

# creates subfolders for each student with files in the correspondent assignment 
# returns a list of students
def create_student_folders( course_path, assignment_path )
	students = []
	#create folders single threaded to prevent race conditions
	@files.each do |file|
		student = file.sharing_user ? file.sharing_user.display_name : file.owners.first.display_name
		student.unicode_normalize!
		user_path = File.join(course_path, assignment_path, student )
		FileUtils.mkdir_p(user_path) unless Dir.exist?(user_path)
		students << student
	end
	return students
end

# will take a list of students (with files in the selected assignment)
# returns a diff with students in @selected_course, i.e. students missing files in the selected assignment
# make sure COURSES_STUDENTS_FILE contains records of all of your courses with their respective students
# returns an empty list [] if COURSES_STUDENTS_FILE is missing or no students are missing files
def get_students_missing_assignment( students_with_assignments )
	students_missing_assignments = []
	begin
		# yaml holds list of students belonging to a specific course
		student_records = YAML.load( File.read( COURSES_STUDENTS_FILE ) )
		students_in_course = student_records[@selected_course[:name]]
		# fix for strings with umlaut (åäö) with ambiguous unicode code point
		students_in_course.map(&:unicode_normalize!) if students_in_course
		
		# this will produce a list of students with empty assignments (no files)
		students_missing_assignments = students_in_course - students_with_assignments if students_in_course
	rescue Errno::ENOENT => e
		puts e
		puts "The configuration file #{COURSES_STUDENTS_FILE} containing a list"
		puts "of your courses with their correspondent students seems to be missing."
	end
	return students_missing_assignments
end

def download_files
	file_count = @files ? @files.length : 0
	puts "\nThere are #{file_count} files in this assignment."
	
	return if !@files
	
	# create folders and subfolders for selected assignment and students
	course_path = @selected_course[:name]
	assignment_path = @selected_assignment[:name]
	FileUtils.mkdir(assignment_path) unless Dir.exist?(assignment_path)
	students_with_assignments = create_student_folders( course_path, assignment_path )

	# the following lines will check for and display students with missing files in the selected assignment
	students_missing_assignments = get_students_missing_assignment( students_with_assignments )
	puts "\nMissing assignments from the following student(s):" if students_missing_assignments.length > 0
	students_missing_assignments.each do |student|
		puts student
	end

	total_file_count = 0
	new_file_count = 0
	mutex = Mutex.new
	new_files = []
	files_checksums = []
	
	puts "\nDownloading:"
	#download in parallel (4 threads in original code)
	Parallel.map(@files, in_threads: 4) do |file|
		# The following properties has no apparent relation to whether file is binary or a google document
		if file.sharing_user
			user_path = File.join(course_path, assignment_path, file.sharing_user.display_name)
		else
			user_path = File.join(course_path, assignment_path, file.owners.first.display_name)
		end
		
		begin
			# normal file (referred to as "binary file" in API reference)
			if file.original_filename
				file_path = get_safe_filepath( file.original_filename, user_path )
				# rename existing copy of file per "basename (yyyy-mm-dd hh.mm.ss).extension"
				file_path_previous_copy = File.exist?( file_path ) ? rename_existing_file( file_path ) : nil
				puts file_path
				@drive_service.get_file(file.id, download_dest: file_path)
			# google document -> export to pdf
			else
				# suffix filename with mime_type (google documents with identical basenames might produce clashing file names)
				mime_type = file.mime_type.to_s.split('.').last		# "application/vnd.google-apps.document" -> "document"
				file_path = get_safe_filepath( "#{file.name} (#{mime_type}).pdf", user_path )
				file_path_previous_copy = File.exist?( file_path ) ? rename_existing_file( file_path ) : nil
				puts file_path
				@drive_service.export_file(file.id, 'application/pdf', download_dest: file_path)
			end
		rescue => e
			puts "\nThere was an error transferring the following file:"
			puts file_path
			puts "Reason: #{e}"
			puts "You might want to try downloading this assignment again.\n"
		end

		mutex.synchronize do
			total_file_count += 1
			files_checksums << { :md5_checksum => file.md5_checksum, :file_path => file_path } if file.md5_checksum
		end
		restore_file_time( file, file_path )		# restore file's mtime from drive
		
		unless remove_duplicate_file_by_timestamp( file_path, file_path_previous_copy )
			mutex.synchronize do
				new_file_count += 1			# update new_file_count if file appears to have no previous copy
				new_files << file_path
			end
		end
	end
	print "\nDownloaded #{total_file_count} file(s)."
	if new_file_count > 0
		print " #{new_file_count} file(s) appear to be modified since last sync:\n"
		new_files.each do |new_file|
			puts new_file
		end
	end
	# single out files with matching md5-checksums:
	duplicate_files = files_checksums.group_by { |h| h[:md5_checksum] }.values.select { |files_checksums| files_checksums.size > 1 }.flatten
	puts "\nThe following files have matching md5-checksums:" if duplicate_files.length > 0
	duplicate_files.each do |hash|
		hash.each { puts "#{hash.values[0]}\t#{hash.values[1]}" }
	end
end

private def initialize_service
	service = Google::Apis::ClassroomV1::ClassroomService.new
	service.client_options.application_name = APPLICATION_NAME
	service.authorization = authorize_service
	return service
end

private def authorize_service
	FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))
	client_id = Google::Auth::ClientId.from_file(CLIENT_SECRETS_PATH)
	token_store = Google::Auth::Stores::FileTokenStore.new(file: CREDENTIALS_PATH)
	authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store)
	user_id = 'default'
	credentials = authorizer.get_credentials(user_id)
	if credentials.nil?
		url = authorizer.get_authorization_url(base_url: OOB_URI)
		puts "Open the following URL in the browser and enter the resulting code after authorization"
		puts url
		code = gets
		credentials = authorizer.get_and_store_credentials_from_code(user_id: user_id, code: code, base_url: OOB_URI)
	end
	credentials
end
end

g = GoogleClassRoomAssignmentDownloader.new