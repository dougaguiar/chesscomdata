# Chess.com Club Data Analyzer

A Python script that uses the official Chess.com API to fetch, analyze, and report on the game activity and rating statistics of all members in a specified chess club. The script generates a detailed multi-tab Excel report, providing valuable insights into player activity, performance, and popular openings.

## Key Features

-   **Comprehensive Data Fetching**: Retrieves all games for every club member within a specified date range.
-   **Performance Analysis**: Calculates initial, final, and overall rating changes (increment) for each player in both Blitz and Rapid categories.
-   **Activity Tracking**: Ranks the top 10 most active players based on the total number of games played.
-   **Opening Repertoire Analysis**:
    -   Identifies the most-played opening for White and defense for Black for each individual player.
    -   Aggregates and reports on the top 10 most popular openings and defenses across the entire club.
-   **Dynamic Summary Report**: Generates a summary tab with key club metrics, including:
    -   Total games played and members processed.
    -   Average club ratings (both at the end of the period and current).
    -   Total rating change for the entire club.
-   **Data Visualization**: Creates a histogram showing the current rating distribution of club members.
-   **Robust and Resilient**: Saves an updated Excel report after each member is processed, preventing data loss from network interruptions.

## Output Example

The script generates a single, formatted Excel file (`club_stats_[club_name]_[timestamp].xlsx`) with the following tabs:

-   **Summary**: High-level club statistics and the rating distribution chart.
-   **Members Raw Data**: A detailed row for each member showing their game counts, rating changes, and preferred openings.
-   **Top 10 Active**: A leaderboard of the most active players.
-   **Popular Openings**: A club-wide ranking of the most frequently played openings with White.
-   **Popular Defenses**: A club-wide ranking of the most frequently played defenses with Black.

## Getting Started

Follow these instructions to get a copy of the project up and running on your local machine.

### Prerequisites

-   Python 3.6 or higher
-   pip (Python package installer)

### Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/dougaguiar/chesscomdata.git
    cd chesscomdata
    ```

2.  **Create and install dependencies:**
    It is highly recommended to use a virtual environment.

    ```bash
    # Create and activate a virtual environment (optional but recommended)
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`

    # Install the required packages
    pip install requests pandas matplotlib openpyxl
    ```

### Script Configuration

Before running the script, you must configure the variables in the `--- CONFIGURATION ---` section at the top of the Python file (`chess_club_analyzer.py`):

-   `CLUB_NAME`: The name of the club as it appears in the chess.com URL (e.g., `torneios-cursos-gm-krikor`).
-   `HEADERS`: **Important:** Replace `'your_username@example.com'` with your own email address. This is recommended by the Chess.com API policy.
-   `YEARS`: Set the start and end year for the analysis. For a single year, set both to the same value (e.g., `[2024, 2024]`).
-   `FIRST_MONTH` / `LAST_MONTH`: The month range for the analysis period.
-   `MEMBER_LIMIT`: For testing, you can set this to a small number (e.g., `5`) to only process the first few members. Set to `None` to process all members.

## Usage

Once the script is configured, run it from your terminal:

```bash
python chess_club_data.py