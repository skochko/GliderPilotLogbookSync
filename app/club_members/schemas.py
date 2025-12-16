from pydantic import BaseModel


class ClubMemberSchema(BaseModel):
    club_id: int
    name: str
    spreadsheet_key: str
    sync_count: int = 0