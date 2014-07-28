CREATE TABLE [dbo].[TrackedFiles] (
    [ID]          INT             IDENTITY (1, 1) NOT NULL,
    [FileName]    NVARCHAR (MAX)  NOT NULL,
    [Tracked]     BIT             DEFAULT 0 NOT NULL,
    [ProjectName] NVARCHAR (1000) NOT NULL,
    PRIMARY KEY CLUSTERED ([ID] ASC)
);


GO
CREATE UNIQUE NONCLUSTERED INDEX [IX_TrackedFiles_Column]
    ON [dbo].[TrackedFiles]([ProjectName] ASC);

