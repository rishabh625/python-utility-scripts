import requests
import pandas as pd
import os

# Replace with your Slack bot token (or use an environment variable)
SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN")  # Or hardcode, but not recommended

# Replace with your channel ID and thread timestamp
CHANNEL_ID = os.environ.get("CHANNEL_ID")
THREAD_TS = os.environ.get("THREAD_TS")
LIMIT = 1000

def get_thread_reactions(channel_id, thread_ts, token, limit=1000):
    """Fetches reactions from a Slack thread and returns data for Excel."""

    url = f"https://slack.com/api/conversations.replies?channel={channel_id}&ts={thread_ts}&limit={limit}"
    headers = {"Authorization": f"Bearer {token}"}

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Raise HTTPError for bad responses (4xx or 5xx)
        data = response.json()

        if not data["ok"]:
            print(f"Error fetching thread replies: {data['error']}")
            return None

        messages = data["messages"]
        reaction_data = {}  # Store aggregated reaction data

        for message in messages:
            user = message.get("user")
            files = message.get("files", [])
            url_private = files[0].get("url_private") if files else None  # handles no files.
            reactions = message.get("reactions", [])
            message_ts = message.get("ts")

            if url_private:
                if user not in reaction_data:
                    reaction_data[user] = {}

                if url_private not in reaction_data[user]:
                    reaction_data[user][url_private] = {}

                for reaction in reactions:
                    reaction_name = reaction["name"]
                    reaction_users = reaction["users"]
                    reaction_count = reaction["count"]

                    if reaction_name in reaction_data[user][url_private]:
                        reaction_data[user][url_private][reaction_name]["users"].extend(reaction_users)
                        reaction_data[user][url_private][reaction_name]["count"] += reaction_count
                    else:
                        reaction_data[user][url_private][reaction_name] = {"users": reaction_users, "count": reaction_count}

        # Prepare data for DataFrame
        aggregated_data = [] #corrected line
        for user, urls in reaction_data.items():
            for url, reactions_by_name in urls.items():
                reaction_names = []
                users_list = []
                counts = []

                for reaction_name, values in reactions_by_name.items():
                    reaction_names.append(reaction_name)
                    users_list.append(",".join(values["users"]))
                    counts.append(values["count"])

                thread_link = f"https://infracloud.slack.com/archives/{CHANNEL_ID}/p{THREAD_TS.replace('.', '')}"
                message_link = f"https://infracloud.slack.com/archives/{CHANNEL_ID}/p{message_ts.replace('.', '')}"

                aggregated_data.append([user, url, ",".join(reaction_names), ",".join(users_list), sum(counts), thread_link, message_link])

        return aggregated_data

    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")
        return None
    except KeyError as e:
        print(f"Key error: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

def replace_user_ids_with_names(data, user_mapping_file):
    """Replaces User IDs with Usernames using a mapping from an Excel file."""

    try:
        user_mapping_df = pd.read_excel(user_mapping_file)
        user_mapping = dict(zip(user_mapping_df["User ID"], user_mapping_df["Username"])) # Corrected column names

        replaced_data = []
        for row in data:
            new_row = []
            for item in row:
                if isinstance(item, str) and item in user_mapping:
                    new_row.append(user_mapping[item])
                elif isinstance(item, str) and "," in item:  # handle user lists
                    user_list = item.split(",")
                    new_user_list = [user_mapping.get(u, u) for u in user_list]  # if user not found, keep id
                    new_row.append(",".join(new_user_list))
                else:
                    new_row.append(item)
            replaced_data.append(new_row)
        return replaced_data

    except FileNotFoundError:
        print(f"Error: User mapping file '{user_mapping_file}' not found.")
        return data  # Return original data if mapping file is not found
    except Exception as e:
        print(f"Error processing user mapping: {e}")
        return data

def replace_users_in_excel(excel_file, user_mapping_file):
    """Reads an Excel file, replaces User IDs in column D, and saves the changes."""

    try:
        df = pd.read_excel(excel_file)
        user_mapping_df = pd.read_excel(user_mapping_file)
        user_mapping = dict(zip(user_mapping_df["User ID"], user_mapping_df["Username"])) # Corrected column names

        def replace_users(users_str):
            if isinstance(users_str, str):
                users = users_str.split(",")
                replaced_users = [user_mapping.get(user, user) for user in users]
                return ",".join(replaced_users)
            return users_str

        df["Users"] = df["Users"].apply(replace_users)
        df.to_excel(excel_file, index=False)
        print(f"User IDs in '{excel_file}' column D replaced with usernames.")

    except FileNotFoundError:
        print(f"Error: Excel file '{excel_file}' or user mapping file '{user_mapping_file}' not found.")
    except Exception as e:
        print(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    if not SLACK_BOT_TOKEN:
        print("Error: Slack bot token must be set as an environment variable.")
    else:

        reactions_data = get_thread_reactions(CHANNEL_ID, THREAD_TS, SLACK_BOT_TOKEN, LIMIT)

        if reactions_data:
            replaced_reactions_data = replace_user_ids_with_names(reactions_data, "slack_users.xlsx")

            df = pd.DataFrame(replaced_reactions_data, columns=["User", "URL", "Reaction Names", "Users", "Total Reaction Count", "Thread Link", "Message Link"])
            output_excel_file = "slack_thread_reactions_messages.xlsx"
            df.to_excel(output_excel_file, index=False)
            print(f"Excel file '{output_excel_file}' created successfully.")

            replace_users_in_excel(output_excel_file, "slack_users.xlsx")
