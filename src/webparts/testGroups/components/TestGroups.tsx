import * as React from 'react'; 
import { useState, useEffect } from 'react';
import styles from './TestGroups.module.scss';
import type { ITestGroupsProps } from './ITestGroupsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';

interface ITag {
  id: string;
  displayName: string;
}

interface IMember {
  id: string;
  displayName: string;
}

// Munkieman Team
//const teamID = "5d3e8ded-4c9f-4bdc-919f-a34ce322caeb";
//const teamName = "TestChat";

//Max Dev Team
const teamID = "696dfe67-e76f-4bf8-8ab6-8abfcb16552e";
const teamName = "TestChat";  

// Max Prod Team
//https://teams.microsoft.com/l/team/19%3AwREFwWCHiIj-qfeAUqedf6wIatZTFqg0CgOwMN6CQxc1%40thread.tacv2/conversations?groupId=a3cce0fc-52f7-4928-8f2b-14102e5ad6ca&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb
//https://teams.microsoft.com/l/channel/19%3AwREFwWCHiIj-qfeAUqedf6wIatZTFqg0CgOwMN6CQxc1%40thread.tacv2/General?groupId=a3cce0fc-52f7-4928-8f2b-14102e5ad6ca&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb
//https://teams.microsoft.com/l/channel/19%3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1%40thread.tacv2/General?groupId=68d9eb2c-06f7-40ed-bd99-a5a35fab0275&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb
//const teamID = "68d9eb2c-06f7-40ed-bd99-a5a35fab0275";
//const teamName = "Teams Testing";  

const channelName = "General";

const TestGroups: React.FunctionComponent<ITestGroupsProps> = (props) => {
  const {
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName,
    context
  } = props;

  const [tags, setTags] = useState<ITag[]>([]);
  const [members, setMembers] = useState<IMember[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  const userEmail = props.context.pageContext.user.email;

  const getTeamTags = (): void => {
    context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3): void => {
        client
          .api(`/teams/${teamID}/tags`)
          .version('v1.0')
          .get((error, response: any) => {
            if (error) {
              console.error('Error fetching tags:', error);
              return;
            }
            setTags(response.value);
          });
      });
      return;
  };

  const fetchChannelMembers = async () : Promise<void> => {
    try {
      setLoading(true);

      // Get Microsoft Graph API client
      const client = await context.msGraphClientFactory.getClient('3');

      // Get joined teams
      const teamsResponse = await client.api('/me/joinedTeams')
        .version('v1.0')
        .get();

      if (!teamsResponse) throw new Error("Failed to fetch teams");

      const team = teamsResponse.value.find((t: any) => t.displayName === teamName);
      if (!team) throw new Error(`Team "${teamName}" not found`);

      // Get channels in the team
      const channelsResponse = await client.api(`/teams/${team.id}/channels`)
        .version('v1.0')
        .get();

      if (!channelsResponse) throw new Error("Failed to fetch channels");

      const channel = channelsResponse.value.find((c: any) => c.displayName === channelName);
      if (!channel) throw new Error(`Channel "${channelName}" not found`);

      // Get channel members
      const membersResponse = await client.api(`/teams/${team.id}/members`)
        .version('v1.0')
        .get();

      if (!membersResponse) throw new Error("Failed to fetch channel members");

      setMembers(membersResponse.value);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
    return;
  };

  const sendMessageToTeams = async (message: string) : Promise<void> => {
    try {
      const client = await context.msGraphClientFactory.getClient('3');
  
      // Fetch Team ID
      const teamsResponse = await client.api('/me/joinedTeams')
        .version('v1.0')
        .get();
      if (!teamsResponse) throw new Error("Failed to fetch teams");
  
      const teamsData = teamsResponse;
      const team = teamsData.value.find((t: any) => t.displayName === teamName);
      if (!team) throw new Error(`Team "${teamName}" not found`);
  
      // Fetch Channel ID
      const channelsResponse = await client.api(`/teams/${team.id}/channels`)
        .version('v1.0')
        .get();
      if (!channelsResponse) throw new Error("Failed to fetch channels");
  
      const channelsData = channelsResponse;
      const channel = channelsData.value.find((c: any) => c.displayName === channelName);
      if (!channel) throw new Error(`Channel "${channelName}" not found`);

      console.log("sendmsg Team:", team);
      console.log("sendmsg Channel:", channel);

      // Fetch Team Tags (To get @expenses Tag ID)
      const tagsResponse = await client.api(`/teams/${team.id}/tags`)
        .version('v1.0')
        .get();
      if (!tagsResponse) throw new Error("Failed to fetch tags");

      const tagsData = tagsResponse;
      const expensesTag = tagsData.value.find((tag: any) => tag.displayName === "expenses");
  
      if (!expensesTag) throw new Error(`Tag "@expenses" not found in team "${teamName}"`);
  
      // 🔥 POST request to send message with @expenses mention
      const mentionId = 1; // You can keep this as 0 or another unique identifier, but it must match the ID in the <at> tag.
      const tagHTML = "<at id='1'>expenses</at> ";

      const response = await client.api(`/teams/${team.id}/channels/${channel.id}/messages`)
        .version('v1.0')
        .post({
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            body: {
              contentType: "html",
              content: tagHTML+message,
            },
            mentions: [
              {
                id: mentionId,
                mentionText: "expenses",
                mentioned: {
                  tag: {
                    id: expensesTag.id,
                    displayName: "expenses",
                  },
                },
              },
            ],
          }),
        });

      console.log("Mention ID:", mentionId);
      console.log("Expenses Tag ID:", expensesTag.id);

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to send message: ${errorText}`);
      }  
    } catch (error: any) {
      console.error("Error sending message:", error.message);
    }
    return;
  };

  const checkMember = async () : Promise<void> => { 
    try {
      const client = await context.msGraphClientFactory.getClient('3');
      const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
      const token = await tokenProvider.getToken("https://graph.microsoft.com");
      console.log("Access Token:", token);

      const userResponse = await client.api(`/users/${userEmail}`)
        .version('v1.0')
        .get();
      const userData = userResponse;
      const userId = userData.id;

      console.log("userID", userId, userData);

      const testResponse = await client.api('/me/joinedTeams')
        .version('v1.0')
        .get();
      const joinedTeams = testResponse;
      console.log("joinedTeams", joinedTeams);

      // Fetch Team ID
      const teamsResponse = await client.api('/me/joinedTeams')
        .version('v1.0')
        .get();
      if (!teamsResponse) throw new Error("Failed to fetch teams");

      const teamsData = teamsResponse;
      let team = teamsData.value.find((t: any) => t.displayName === teamName);

      if (!team) {
        console.log(`User is not in the team "${teamName}". Fetching team ID manually...`);

        // Get all teams the user has access to
        const allTeamsResponse = await client.api(`/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')`)
          .version('v1.0')
          .get();
        if (!allTeamsResponse) throw new Error("Failed to fetch available teams");

        const allTeamsData = allTeamsResponse;
        team = allTeamsData.value.find((t: any) => t.displayName === teamName);

        if (!team) throw new Error(`Team "${teamName}" not found.`);
      }

      console.log("Found Team:", team);
      console.log("checkmember Team:", team);

      // Check if user is already a member of the team
      const membersResponse = await client.api(`/teams/${team.id}/members`)
        .version('v1.0')
        .get();
      const membersData = membersResponse;
      const userIsMember = membersData.value.some((m: any) => m.id === userId);

      console.log("Found Team:", team);
      console.log("checkmember Team:", team);
      console.log("useIsMember", userIsMember);

      if (!userIsMember) {
        console.log("User is not a member, adding to the chat channel...");

        // Add user to the team
        const addUserResponse = await client.api(`/teams/${team.id}/members`)
          .version('v1.0')
          .post({
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              "roles": ["member"],
              "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userId}`,
              "isHistoryIncluded": false,
              "visibleHistoryStartDateTime": null
            })
          });

        if (!addUserResponse.ok) {
          const errorText = await addUserResponse.text();
          throw new Error(`Failed to add user to the chat: ${errorText}`);
        }

        console.log("User successfully added to the chat");
      } else {
        console.log("User is already a member of the chat channel");
      }

      //fetchTeamMembers();

    } catch (error: any) {
      console.error("Error in checkMember:", error.msg);
    }
    return;
  };

  useEffect(() => {    
    fetchChannelMembers();
    checkMember();
    getTeamTags();
    console.log("Member check completed."); 
      
      // 🔥 Post message to Teams channel
      setTimeout(async() => {
        await sendMessageToTeams("this is a test message");
      }, 3000);          
  }, []);


  return (
    <section className={`${styles.testGroups} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
      </div>
      <div>
        <h2>Team Tags</h2>
        <ul>
          {tags.map(tag => (
            <li key={tag.id}>{tag.displayName}</li>
          ))}
        </ul>
      </div>
      <div>
        <h2>Channel Members</h2>
        {loading && <p>Loading members...</p>}
        {error && <p>Error: {error}</p>}
        <ul>
          {members.map(member => (
            <li key={member.id}>{member.displayName}</li>
          ))}
        </ul>
      </div>
    </section>
  );
};

export default TestGroups;