## [PowerPoint MCP Agent](https://github.com/Coral-Protocol/Coral-PowerPointMCP-Agent)

PowerPoint agent capable of creating comprehensive presentations with professional design, intelligent content organization, and advanced PowerPoint manipulation features using MCP tools.

Link- [PowerPoint MCP Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server)

## Responsibility
PowerPoint MCP Agent is a comprehensive presentation creation system that leverages the Model Context Protocol (MCP) to provide AI agents with advanced PowerPoint manipulation capabilities. It offers 32+ specialized tools for creating professional presentations with dynamic content, intelligent layout selection, advanced formatting, and seamless integration with multi-agent workflows through the Coral Protocol.

## Details
- **Framework**: Pydantic Ai
- **Tools used**: Office PowerPoint MCP Server Tools (32 tools), Coral Server Tools
- **AI model**: OpenAI GPT-4o, Groq, or any compatible LLM provider
- **Date added**: December 2024
- **Reference**: [Office PowerPoint MCP Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server)
- **License**: MIT

## Setup the Agent

### 1. Clone & Install Dependencies

<details>

Ensure that the [Coral Server](https://github.com/Coral-Protocol/coral-server) is running on your system. If you are trying to run the PowerPoint agent and require an input, you can either create your agent which communicates on the coral server or run and register the [Interface Agent](https://github.com/Coral-Protocol/Coral-Interface-Agent) on the Coral Server.

```bash
# In a new terminal clone the repository:
git clone https://github.com/Coral-Protocol/Coral-PowerPointMCP-Agent

# Navigate to the project directory:
cd Coral-PowerPointMCP-Agent

# install uv
pip install uv

# Install dependencies from `pyproject.toml` using `uv`:
uv sync
```

</details>

### 2. Configure Environment Variables

<details>

Get the API Key:
[OpenAI](https://platform.openai.com/api-keys) || 
[Groq](https://console.groq.com/keys) || 
[Anthropic](https://console.anthropic.com/) || 
Any other compatible LLM provider

```bash
# Create .env file in project root
cp -r .env_sample .env
```

Configure your .env file with the following variables:
```bash
# LLM Configuration
MODEL_PROVIDER=openai  # or groq, anthropic, etc.
MODEL_NAME=gpt-4o      # or moonshotai/kimi-k2-instruct for Groq
OPENAI_API_KEY/GROQ_API_KEY =your_api_key_here

```
Check if the .env file has correct URL for Coral Server and adjust the parameters accordingly.

</details>

### 3. Configure Logfire for Monitoring

<details>

This agent uses [Logfire](https://ai.pydantic.dev/logfire/#pydantic-logfire) for monitoring tool usage and agent behavior. Logfire provides detailed insights into how your agent interacts with tools and handles requests.

1. **Setup Logfire**
   ```python
   # These lines are already included in main.py
   import logfire
   
   logfire.configure()  
   logfire.instrument_pydantic_ai()
   ```

For more details about Logfire configuration and features, visit: [Pydantic-AI Logfire Documentation](https://ai.pydantic.dev/logfire/#pydantic-logfire)

</details>

## Run the Agent

You can run in either of the below modes to get your system running.

- The Executable Model is part of the Coral Protocol Orchestrator which works with [Coral Studio UI](https://github.com/Coral-Protocol/coral-studio).
- The Dev Mode allows the Coral Server and all agents to be separately running on each terminal without UI support.

### 1. Executable Mode

Checkout: [How to Build a Multi-Agent System with Awesome Open Source Agents using Coral Protocol](https://github.com/Coral-Protocol/existing-agent-sessions-tutorial-private-temp) and update the file: `coral-server/src/main/resources/application.yaml` with the details below, then run the [Coral Server](https://github.com/Coral-Protocol/coral-server) and [Coral Studio UI](https://github.com/Coral-Protocol/coral-studio). You do not need to set up the `.env` in the project directory for running in this mode; it will be captured through the variables below.

<details>

For Linux or MAC:

```bash
# PROJECT_DIR="/PATH/TO/YOUR/PROJECT"

applications:
  - id: "app"
    name: "Default Application"
    description: "Default application for testing"
    privacyKeys:
      - "default-key"
      - "public"
      - "priv"

registry:
  powerpoint_agent:
    options:
      - name: "GROQ_API_KEY"
        type: "string"
        description: "API key for the LLM service (OpenAI, Groq, etc.)"
      - name: "MODEL_PROVIDER"
        type: "string"
        description: "LLM provider (openai, groq, anthropic, etc.)"
      - name: "MODEL_NAME"
        type: "string"
        description: "Model name (gpt-4o, llama-3.3-70b-versatile, etc.)"
    runtime:
      type: "executable"
      command: ["bash", "-c", "${PROJECT_DIR}/run_agent.sh main.py"]
      environment:
        - name: "GROQ_API_KEY"
          from: "GROQ_API_KEY"
        - name: "MODEL_PROVIDER"
          from: "MODEL_PROVIDER"
        - name: "MODEL_NAME"
          from: "MODEL_NAME"
        - name: "MODEL_TEMPERATURE"
          value: "0.1"
        - name: "MODEL_TOKEN"
          value: "8000"
```

For Windows, create a powershell command (run_agent.ps1) and run:

```bash
command: ["powershell","-ExecutionPolicy", "Bypass", "-File", "${PROJECT_DIR}/run_agent.ps1","main.py"]
```

</details>

### 2. Dev Mode

Ensure that the [Coral Server](https://github.com/Coral-Protocol/coral-server) is running on your system and run below command in a separate terminal.

<details>

```bash
# Run the agent using `uv`:
uv run main.py
```

You can view the agents running in Dev Mode using the [Coral Studio UI](https://github.com/Coral-Protocol/coral-studio) by running it separately in a new terminal.

</details>

## Example

<details>

```bash
# Input to the interface agent:
Ask the Office PowerPoint MCP Agent to generate a PowerPoint presentation on the topic "Transformer Architecture". The agent should analyze the topic and create a clean, professional, and visually appealing deck that includes relevant tables, charts, and shapes. Ensure all text is properly contained within the slide boundaries for a neat and tidy appearance. Save the final presentation as "transformer_architecture" at the specified path.

# Output:
It will create the slides at the path you provided.
```

</details>

### Creator Details
- **Name**:Ahsen Tahir
- **Affiliation**: Coral Protocol
- **Contact**: [Discord](https://discord.com/invite/Xjm892dtt3)
