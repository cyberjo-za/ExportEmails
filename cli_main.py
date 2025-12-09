import aspose.email as ae
from getpass import getpass
import os
import json
from datetime import datetime

SESSION_FILE = "export_session.json"

def load_session():
    """Load previous session progress."""
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE, 'r') as f:
                return json.load(f)
        except:
            return None
    return None

def save_session(session_data):
    """Save session progress."""
    with open(SESSION_FILE, 'w') as f:
        json.dump(session_data, f, indent=2)
    print(f"✓ Session saved: {session_data['processed_count']} emails exported")

def export_imap_to_pst(server, port, email, password, output_file):
    """Connect to IMAP server and export to PST file with progress tracking."""
    try:
        # Create IMAP client
        imap_client = ae.clients.imap.ImapClient()
        imap_client.host = server
        imap_client.port = port
        imap_client.username = email
        imap_client.password = password
        imap_client.security_options = ae.clients.SecurityOptions.AUTO
        imap_client.timeout = 300000  # 5 minutes timeout
        
        print(f"✓ Connected to {server}")
        
        # Get mailbox info
        mailbox_info = imap_client.mailbox_info
        print(f"✓ Retrieved mailbox info")
        
        # Get inbox folder info
        inbox_name = mailbox_info.inbox.name
        inbox_info = imap_client.get_folder_info(inbox_name)
        print(f"✓ Folder: {inbox_name}")
        
        # Create PST file
        pst = ae.storage.pst.PersonalStorage.create(output_file, ae.storage.pst.FileFormatVersion.UNICODE)
        pst_folder = pst.root_folder.add_sub_folder(inbox_name)
        
        # Load previous session
        session = load_session()
        start_index = 0
        if session and session.get('output_file') == output_file:
            start_index = session.get('processed_count', 0)
            print(f"✓ Resuming from email {start_index + 1}")
        
        # Select folder and get messages
        imap_client.select_folder(inbox_name)
        messages = imap_client.list_messages()
        total_messages = len(messages)
        
        print(f"✓ Found {total_messages} messages in {inbox_name}")
        print(f"✓ Starting export (timeout set to 5 minutes per batch)...\n")
        
        # Process messages in batches
        processed = start_index
        batch_size = 10
        
        for i in range(start_index, total_messages):
            try:
                msg_info = messages[i]
                
                # Fetch message from IMAP
                email_msg = imap_client.fetch_message(msg_info.unique_id)
                
                # Convert MailMessage to MapiMessage for PST compatibility
                mapi_msg = ae.mapi.MapiMessage.from_mail_message(email_msg)
                pst_folder.add_message(mapi_msg)
                
                processed += 1
                progress = (processed / total_messages) * 100
                
                # Print progress every 10 messages
                if processed % batch_size == 0 or processed == total_messages:
                    print(f"  Progress: [{processed}/{total_messages}] ({progress:.1f}%) - {msg_info.subject[:50]}")
                    
                    # Save session every batch
                    save_session({
                        'output_file': output_file,
                        'processed_count': processed,
                        'total_count': total_messages,
                        'last_updated': datetime.now().isoformat()
                    })
                
            except Exception as e:
                print(f"  ✗ Failed to export message {i + 1}: {e}")
                # Save progress and continue
                save_session({
                    'output_file': output_file,
                    'processed_count': processed,
                    'total_count': total_messages,
                    'last_error': str(e),
                    'last_updated': datetime.now().isoformat()
                })
                continue
        
        # Save PST file
        pst.save()
        pst.dispose()
        
        print(f"\n✓ Successfully exported {processed} emails to {output_file}")
        
        # Clean up session file on success
        if os.path.exists(SESSION_FILE):
            os.remove(SESSION_FILE)
            print("✓ Session cleared")
        
        return True
        
    except Exception as e:
        print(f"✗ Export failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("=" * 50)
    print("IMAP to PST Exporter (with Resume Support)")
    print("=" * 50)
    
    # Check for existing session
    session = load_session()
    if session:
        print(f"\n✓ Found previous session:")
        print(f"  File: {session.get('output_file')}")
        print(f"  Processed: {session.get('processed_count')}/{session.get('total_count')}")
        resume = input("\nResume previous export? (y/n): ").strip().lower()
        if resume == 'y':
            output_file = session['output_file']
            # Need to get connection details again
            server = input("Enter IMAP server: ").strip()
            port = int(input("Enter port: ").strip())
            email = input("Enter email address: ").strip()
            password = getpass("Enter password: ")
            export_imap_to_pst(server, port, email, password, output_file)
            print("\n✓ Done!")
            return
        else:
            os.remove(SESSION_FILE)
            print("✓ Session cleared\n")
    
    # New export
    print()
    server = input("Enter IMAP server (e.g., imap.gmail.com): ").strip()
    port = int(input("Enter port (usually 993 for SSL): ").strip())
    email = input("Enter email address: ").strip()
    password = getpass("Enter password: ")
    
    # Get output filename
    output_file = input("\nEnter output PST filename (e.g., exported.pst): ").strip()
    if not output_file.endswith('.pst'):
        output_file += '.pst'
    
    # Check if file already exists
    if os.path.exists(output_file):
        overwrite = input(f"File '{output_file}' already exists. Overwrite? (y/n): ").strip().lower()
        if overwrite != 'y':
            print("✗ Cancelled")
            return
        os.remove(output_file)
    
    # Export
    export_imap_to_pst(server, port, email, password, output_file)
    print("\n✓ Done!")

if __name__ == "__main__":
    main()