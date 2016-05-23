# encoding: utf-8

namespace :a5 do
  desc 'Parse a5'
  task parse_tasks: :environment do
    require 'open-uri'

    CREDENTIALS      = %w(7pikes password)
    BASE_URL         = 'http://help.apteka5.ru'
    TASK_URL         = BASE_URL + '/Task/View/'
    TASK_API_URL     = BASE_URL + '/api/task/'
    OUR_FILTER_URL   = TASK_API_URL + '?pagesize=100&filterid=' # pagesize=100(maximum),т.к. default = 25
    # 250 => 7 pikes (Текущий), 975  => __ИНВЕНТЫ! МСП!!!, 1132 => 7 pikes (Глобальные), 1112 => 7 pikes (Отложено)
    FILTERS_TO_PARSE = [250, 975, 1132]

    close_issues_in_db
    FILTERS_TO_PARSE.each { |filter| parse_by_filter(filter) }
    close_issues
  end

  private

  def close_issues_in_db
    Issue.update_all(is_opened_on_helpdesk: false)
  end

  def parse_by_filter(filter)
    filter_url = OUR_FILTER_URL + filter.to_s
    opened_tasks_ids = get_opened_tasks(filter_url)

    puts 'Adding new issues, updating comments, saving attachments'
    opened_tasks_ids.each do |a5_task_id|
      issue = Issue.find_by_a5_helpdesk_id(a5_task_id)
      if issue.blank?
        create_issue(a5_task_id)
      else
        update_issue(a5_task_id, issue)
      end
    end
  end

  def update_issue(a5_task_id, issue)
    puts "Updating issue #{ a5_task_id }"

    reopen_issue(issue)

    db_comment_ids        = issue.comments.pluck(:comment_id)
    db_attachment_ids     = issue.attachments.pluck(:attachment_id)

    helpdesk_attachments  = get_attachments(a5_task_id)
    helpdesk_comments     = get_comments(a5_task_id)

    help_attachment_ids   = helpdesk_attachments.values.map { |value| value.split('/').last.to_i }
    help_comment_ids      = helpdesk_comments.map { |value| value[:pid].to_i }

    new_attachments       = help_attachment_ids - db_attachment_ids
    new_comments          = help_comment_ids    - db_comment_ids

    comments              = helpdesk_comments.reject { |comment| db_comment_ids.include?(comment[:pid].to_i) }

    update_comments_and_attachments(a5_task_id, new_attachments, comments, issue) unless (new_comments || new_attachments).blank?
    done
  end

  def update_comments_and_attachments(a5_task_id, new_attachments, comments, issue)
    print "Updating attachments and comments for a5 task ##{ a5_task_id }"

    files = {}
    task_url   = TASK_API_URL + "#{ a5_task_id }"
    task_xml   = get_xml(task_url)
    file_ids   = task_xml.xpath('//FileIds').text.split(',')
    file_names = task_xml.xpath('//Files').text.split(',')
    unless file_names.blank?
      new_attachments.each do |id|
        file_name        = file_names[file_ids.index(id.to_s)]
        files[file_name] = "http://help.apteka5.ru/api/taskfile/#{ id }"
      end
    end

    # preparing hashes of attachments
    hash_of_attachments = []
    attachments = create_attachments(files)
    attachments.each do |attachment|
      hash_of_attachments << { 'hash' => attachment }
    end

    # saving attachments to database
    files.values.each do |link|
      Attachment.create(attachment_id: link.split('/').last.to_i, issue_id: issue.id)
    end

    # creating threads
    unless hash_of_attachments.blank?
      print 'Creating thread with new attachments'
      thread_with_attachments = {
        'type'        => 'customer',
        'createdBy'   => { 'id' => 48634526, 'type' => 'customer' },
        'status'      => 'active',
        'attachments' => hash_of_attachments,
        'body'        => 'В заявке обновились прикрепленные файлы:'
      }

      thread = HelpScout::Conversation::Thread.new(thread_with_attachments)
      HelpScout::Client.new.create_thread(issue.helpscout_id, thread)
    end

    unless comments.blank?
      print 'Creating threads with new comments'
      comments.each do |comment|
        created_at = comment[:created_at].split(' ').first + ' в ' + comment[:created_at].split(' ').last
        comment_body = "<b>#{ comment[:editor] }</b>, #{ created_at }: \n #{ comment[:comment] }"
        thread = HelpScout::Conversation::Thread.new({
          'type'      => 'customer',
          'createdBy' => { 'id' => 48634526, 'type' => 'customer' },
          'body'      => comment_body,
          'status'    => 'active'
        })
        HelpScout::Client.new.create_thread(issue.helpscout_id, thread)
        Comment.create(issue_id: issue.id, comment_id: comment[:pid])
      end
    end
  end

  def reopen_issue(issue)
    puts 'Reopening issue in db and if necessary on helpscout server'
    issue.update_column(:is_opened_on_helpdesk, true)
    if issue.is_opened_on_helpscout == false
      puts 'Issue was closed on helpscout, reopening'
      conversation = HelpScout::Conversation.new('id' => issue.helpscout_id, 'status' => 'active')
      helpscout_response = HelpScout::Client.new.update_conversation(conversation)
      issue.update_column(:is_opened_on_helpscout, true) if helpscout_response == 200
    end
  end

  def get_opened_tasks(filter_url)
    print 'Getting list of opened tasks'
    opened_tasks_ids = []
    first_page_xml = get_xml(filter_url)
    first_page_xml.xpath('//Id').each { |id| opened_tasks_ids << id.text }
    page_count = first_page_xml.xpath('//PageCount').text.to_i
    if page_count > 1
      (2..page_count).each do |page|
        all_tasks_xml = get_xml(filter_url + "&page=#{ page }")
        all_tasks_xml.xpath('//Id').each { |id| opened_tasks_ids << id.text }
      end
    end

    done
    puts "There are #{ opened_tasks_ids.count } opened tasks in filter #{ filter_url.to_s }"
    opened_tasks_ids.reverse
  end

  def create_issue(a5_task_id)
    puts "Creating issue on helpscout server"
    task_url  = TASK_API_URL + a5_task_id.to_s
    task_xml  = get_xml(task_url)
    executors = task_xml.xpath('//Executors').text
    if executors.include?('7 Pikes')
      conversation       = new_conversation(task_xml, a5_task_id)
      helpscout_response = HelpScout::Client.new.create_conversation(conversation)
      helpscout_id       = helpscout_response[/\d+\d/] if helpscout_response.include?('https')
      if helpscout_id
        save_everything_to_db(helpscout_id, a5_task_id)
        leave_comment(helpscout_id, a5_task_id)
      end
    end
    done
  end

  def save_everything_to_db(helpscout_id, a5_task_id)
    print 'Saving everything to db'
    issue = Issue.create(
      helpscout_id:           helpscout_id,
      a5_helpdesk_id:         a5_task_id,
      is_opened_on_helpdesk:  true,
      is_opened_on_helpscout: true
    )

    attachments = get_attachments(a5_task_id)
    unless attachments.blank?
      attachments.values.each do |link|
        Attachment.create(attachment_id: link.split('/').last.to_i, issue_id: issue.id)
      end
    end

    comments = get_comments(a5_task_id)
    unless comments.blank?
      comments.each do |comment|
        Comment.create(issue_id: issue.id, comment_id: comment[:pid])
      end
    end
    done
  end

  def new_conversation(task_xml, a5_task_id)
    subject     = task_xml.xpath('//Name').text
    priority    = task_xml.xpath('//PriorityName').text
    description = task_xml.xpath('//Description').text
    creator     = task_xml.xpath('//Creator').text
    # deadline    = task_xml.xpath('//Deadline').text
    created_at  = task_xml.xpath('//Created').text
    body        = prepare_body(a5_task_id, creator, description, created_at)
    tags        = prepare_tags(creator, priority, body, description)
    comments    = get_comments(a5_task_id)
    attachments = prepare_attachments(a5_task_id)
    threads     = prepare_threads(comments, body, attachments)

    print "Creating task with subject: #{ subject }. Task on a5 helpdesk: ##{ a5_task_id }"
    conversation = HelpScout::Conversation.new(
      'folderId' => 408173,
      'type'     => 'email',
      'isDraft'  => 'false',
      'status'   => 'active',
      'subject'  => subject,
      'tags'     => tags,
      'mailbox'  => { 'id' => 37534 },
      'customer' => { 'id' => 48634526 },
      'threads'  => threads
    )

    done
    conversation
  end

  def get_xml(url)
    Nokogiri::XML(open(url, http_basic_authentication: CREDENTIALS, 'accept' => 'application/xml'))
  end

  def prepare_threads(comments, body, attachments)
    print 'Preparing threads'
    threads = [
      {
        'type'        => 'customer',
        'createdBy'   => { 'id' => 48634526, 'type' => 'customer' },
        'body'        => body,
        'status'      => 'active',
        'attachments' => attachments
      }
    ]

    comments.each do |comment|
      created_at = comment[:created_at].split(' ').first + ' в ' + comment[:created_at].split(' ').last
      comment_body = " \n \n <b>#{ comment[:editor] }</b>, #{ created_at }: \n #{ comment[:comment] }"
      threads << {
        'type'      => 'customer',
        'createdBy' => { 'id' => 48634526, 'type' => 'customer' },
        'body'      => comment_body,
        'status'    => 'active'
      }
    end
    done
    threads
  end

  def prepare_tags(creator, priority, body, description)
    print 'Preparing tags'
    unless priority == 'критический'
      ['аптека стоит', 'Аптека стоит', 'АПТЕКА СТОИТ'].each do |text|
        priority = 'критический' if body.include?(text) || description.include?(text)
      end
    end

    unless priority == 'критический'
      ['срочно', 'Срочно', 'СРОЧНО'].each do |text|
        priority = 'высокий' if body.include?(text) || description.include?(text)
      end
    end

    tags = ['А5', "#{ priority } приоритет"]
    tags << "МСП, Инвент" if ["Молчанова Юлия", "Козлова Елена"].include?(creator)
    ['Ошибка синхронизации', 'невозможно списать товар', 'МСП'].each do |text|
      tags << "МСП, Инвент" if body.include?(text) || description.include?(text)
    end
    done

    tags
  end

  def get_comments(a5_task_id)
    print "processing comments on task #{a5_task_id}"

    task_comments_url = "http://help.apteka5.ru/api/tasklifetime?taskid=#{ a5_task_id }&include=Comments"
    task_comments_xml = get_xml(task_comments_url)
    task_lifetimes    = task_comments_xml.xpath('//TaskLifetime')
    comments          = task_comments_xml.xpath('//Comments')
    dates             = []
    editors           = []
    task_lifetimes.each do |task_lifetime|
      unless task_lifetime.xpath('Comments').blank?
        dates   << task_lifetime.xpath('Date').text
        editors << task_lifetime.xpath('Editor').text
      end
    end

    result = []
    comments.each do |comment|
      a5_comment_created_on = dates[comments.index(comment)]
      a5_comment            = comment.text
      pid                   = a5_comment_created_on.gsub(/[^0-9]/, '')
      editor                = editors[comments.index(comment)]

      result << { created_at: a5_comment_created_on, editor: editor, comment: a5_comment, pid: pid }
    end
    done
    result
  end

  def prepare_body(a5_task_id, creator, description, created_at)
    print 'Preparing body'
    created_at = created_at.split(' ').first + ' в ' + created_at.split(' ').last
    body = "#{ TASK_URL + a5_task_id.to_s } \n \n <b>#{ creator } создал/a заявку</b> #{ created_at }: \n #{ description }"
    done
    body
  end

  def get_attachments(a5_task_id)
    print "Getting attachments for a5 task ##{ a5_task_id }"
    files = {}
    task_url   = TASK_API_URL + "#{ a5_task_id }"
    task_xml   = get_xml(task_url)
    file_ids   = task_xml.xpath('//FileIds').text.split(',')
    file_names = task_xml.xpath('//Files').text.split(',')
    unless file_names.blank?
      file_ids.each do |id|
        file_name        = file_names[file_ids.index(id)]
        files[file_name] = "http://help.apteka5.ru/api/taskfile/#{ id }"
      end
    end
    done
    files
  end

  def create_attachments(attachments)
    print 'Creating attachments on helpscout server'
    result = []
    attachments.each do |name, link|
      file      = Base64.encode64(open(link, http_basic_authentication: CREDENTIALS).read)
      extension = name.to_s.split('.').last
      mime_type = Mime::Type.lookup_by_extension(extension.downcase).to_s

      attachment = {
        'fileName' => name,
        'mimeType' => mime_type,
        'data'     => file
      }
      result << HelpScout::Client.new.create_attachment(attachment)['hash']
    end
    done
    result
  end

  def prepare_attachments(a5_task_id)
    puts 'Creating attachments hashes'
    result = []
    attachments = get_attachments(a5_task_id)
    attachments = create_attachments(attachments)
    attachments.each do |attachment|
      result << { 'hash' => attachment }
    end
    done
    result
  end

  def close_issues
    puts 'Closing issues that are not in the xml feeds'
    tasks_to_close = Issue.where(is_opened_on_helpscout: true, is_opened_on_helpdesk: false).pluck(:helpscout_id)
    puts "Closing #{ tasks_to_close.count } issues"

    tasks_to_close.each do |helpscout_id|
      issue = Issue.find_by_helpscout_id(helpscout_id)
      puts "Closing task #{ issue.a5_helpdesk_id }"
      conversation = HelpScout::Conversation.new('id' => helpscout_id, 'status' => 'closed')
      helpscout_response = HelpScout::Client.new.update_conversation(conversation)
      issue.update_column(:is_opened_on_helpscout, false) if helpscout_response == 200
    end
    done
  end

  def leave_comment(helpscout_id, a5_task_id)
    puts 'Adding comment to a5 helpdesk'
    conn = Faraday.new(url: BASE_URL)
    conn.basic_auth CREDENTIALS[0], CREDENTIALS[1]

    conn.put do |req|
      req.url "/api/Task/#{ a5_task_id }"
      req.headers['Content-Type'] = 'application/json'
      req.body = "{'Comment': 'Заявка зарегистрирована под номером #{ HelpScout::Client.new.conversation(helpscout_id).number }'}"
    end
    done
  end

  def done
    puts "\033[32m...Done\033[0m"
  end
end


