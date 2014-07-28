CREATE TABLE [dbo].[Metrics] (
    [ID]        INT        IDENTITY (1, 1) NOT NULL,
    [FileID]    INT        NOT NULL,
    [TimeStamp] DATETIME NOT NULL,
    [WordCount] INT        NOT NULL,
    PRIMARY KEY CLUSTERED ([ID] ASC),
    CONSTRAINT [FK_Metrics_ToTable] FOREIGN KEY ([FileID]) REFERENCES [dbo].[TrackedFiles] ([ID])
);

