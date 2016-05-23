# encoding: utf-8

namespace :ringcloud do
  desc 'Parse a5'
  task parse_calls: :environment do
    require 'open-uri'
    HELPSCOUT_CUSTOMER_ID = 55326153
    RINGCLOUD_API_KEY     = 'api_key'
    RINGCLOUD_PASSWORD    = 'password'
    RINGCLOUD_HASH        = 'hash'
    FINISHED_CALLS_URL    = "https://api.ringcloud.ru/v1/calls/complete?api_key=#{ RINGCLOUD_API_KEY }&hash=#{ RINGCLOUD_HASH }&days=1"
    RINGCLOUD_URL         = 'https://api.ringcloud.ru/'

    json = get_json(FINISHED_CALLS_URL)
    json['data'].each do |call|
      conversation       = new_conv(call)
      helpscout_response = HelpScout::Client.new.create_conversation(conversation) unless conversation.nil?
    end
  end

  private

  def new_conv(call)
    phone_number = '+' + call['src']
    subject      = 'Звонок от ' + phone_number
    file_link    = RINGCLOUD_URL + call['rec_file'] + '?api_key=' + RINGCLOUD_API_KEY + '&hash=' + RINGCLOUD_HASH
    file_name    = call['rec_file'].split('/').last
    if Call.find_by_file_name(file_name).nil? && file_name != 'None'
      attachment   = prepare_attachment(file_name, file_link)
      thread       = prepare_thread(phone_number, attachment)

      conversation = HelpScout::Conversation.new(
        'folderId' => 408173,
        'type'     => 'phone',
        'isDraft'  => 'false',
        'status'   => 'active',
        'subject'  => subject,
        'mailbox'  => { 'id' => 37534 },
        'customer' => { 'id' => HELPSCOUT_CUSTOMER_ID },
        'threads'  => thread
      )

      done
      conversation
    else
      nil
    end
  end

  def get_json(url)
    JSON.parse(open(url).read)
  end

  def create_attachment(file_name, file_link)
    print 'Creating attachment on helpscout server'
    result = []
    file   = Base64.encode64(open(file_link).read)

    attachment = {
      'fileName' => file_name,
      'mimeType' => 'audio/mpeg3',
      'data'     => file
    }
    result << HelpScout::Client.new.create_attachment(attachment)['hash']
    
    Call.create(file_name: file_name)
    
    done
    result
  end

  def prepare_attachment(file_name, file_link)
    puts 'Creating attachments hashes'
    result = []
    attachments = create_attachment(file_name, file_link)
    attachments.each do |attachment|
      result << { 'hash' => attachment }
    end

    done
    result
  end

  def prepare_thread(phone_number, attachment)
    print 'Preparing threads'
    thread = [
      {
        'type'        => 'customer',
        'createdBy'   => { 'id' => HELPSCOUT_CUSTOMER_ID, 'type' => 'customer' },
        'body'        => "Прокомментируйте звонок от #{ phone_number }, запись во вложении.",
        'status'      => 'active',
        'attachments' => attachment
      }
    ]

    done
    thread
  end

  def done
    puts "\033[32m...Done\033[0m"
  end
end
